<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosSeguridad.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosSQL.asp"-->
<%
'--------------------------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------------------------
Function getDsEstadoAjuste(pCdEstado)
	Dim str
	select case pCdEstado
		case AJUSTE_ESTADO_AUTORIZADO
			str = "Autorizado"
		case AJUSTE_ESTADO_CANCELADO
			str = "Cancelado"
        case else
			str = "No Autorizado"
	end select
	getDsEstadoAjuste = str
End Function
'----------------------------------------------------------------------------------------------------------------------
Function getAjustesPto(pto, pNajuste, pOrigen, pcdProducto, pkilos, pfechaDesde, pfechaHasta, pEstado, pOrder, pConcepto)
	Dim strSQL, myWhere
	myWhere = filtrarAjustes(pNajuste, pOrigen, pcdProducto, pkilos, pfechaDesde, pfechaHasta, pEstado, pConcepto)
	'NOTA: la SQL esta pensada para buscar archivos de Draft Survey, se planteo asi por que es la unica tabla que trabaja con 
	'los ajustes y tiene para descargar Adjuntos, en caso de agregarse una nueva tabla que tenga archivos adjuntos y
	'se relacione a algun codigo de Ajuste se debera modificar esta SQL
	strSQL = " SELECT A.*, C.DSPRODUCTO, B.NAMEFILE, B.EXTFILE " &_
			 " FROM ( " &_
			 "		 SELECT * FROM TBLAJUSTES " & myWhere &_
			 "		) A " &_
			 " LEFT JOIN TBLEMBARQUESDRAFTSURVEY B ON A.CDAJUSTE = '"& AJUSTE_DRAFT_SURVEY &"' AND B.IDDRAFT = A.IDORIGEN "&_
			 " INNER JOIN  PRODUCTOS C ON C.CDPRODUCTO = A.CDPRODUCTO " & pOrder
	
    call GF_BD_Puertos (pto, rs, "OPEN",strSQL)
	Set getAjustesPto = rs
End Function
'----------------------------------------------------------------------------------------------------------------------
Function filtrarAjustes(pNajuste, pOrigen, pcdProducto, pkilos, pfechaDesde, pfechaHasta, pEstado, pConcepto)
	Dim ret, strCdAjs
    if (CInt(pEstado) <> 0) then
        'Si en el filtro de busqueda selecciono no autorizado se muestra todos los ajustes que esten pendientes y los que se estan firmando    
        if (CInt(pEstado) = AJUSTE_ESTADO_NOAUTORIZADO) then
            ret = ret & "WHERE ESTADO >= "& AJUSTE_ESTADO_NOAUTORIZADO &" AND ESTADO < "& AJUSTE_ESTADO_AUTORIZADO
        else
            ret = ret & "WHERE ESTADO = "& pEstado
        end if
    end if
	if (IsNumeric(pNajuste)) then
		if ((pNajuste <> 0)and(pNajuste <> "")) then Call mkWhere(ret, "IDAJUSTE", pNajuste, "=", 1)
	end if	
	
    if (pConcepto <> "") then Call mkWhere(ret, "CDAJUSTE", pConcepto, "=", 3)   
    
	if (IsNumeric(pOrigen)) then
		if ((pOrigen <> 0)and(pOrigen <> "")) then Call mkWhere(ret, "IDORIGEN", pOrigen, "=", 1)
	end if	
	if (pcdProducto > 0) then  Call mkWhere(ret, "CDPRODUCTO", pcdProducto, "=", 1)
	if (IsNumeric(pkilos)) then
		 if ((pkilos <> 0)and(pkilos <> ""))then Call mkWhere(ret, "KILOSAJUSTE", pkilos, "=", 1)
	end if
	if (pfechaDesde <> "") then  Call mkWhere(ret, "FECHADESDE", pfechaDesde, ">=", 1)
	if (pfechaHasta <> "") then  Call mkWhere(ret, "FECHAHASTA", pfechaHasta, "<=", 1)
	
	filtrarAjustes = ret	
End Function
'----------------------------------------------------------------------------------------------------------------------
Dim Conn,g_strPuerto,mostrar,paginaActual,g_Najuste,g_CdAjuste,g_Origen,g_cdProducto,g_kilos,g_Estado,fecAjusteD,fecAjusteM
Dim fecAjusteA,fecAjusteDH,fecAjusteMH,fecAjusteAH,fechaDesde,fechaHasta, estadoSearch,chk1,chk2,chk3,concepto

g_strPuerto   = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)

Call initTaskAccessInfo(TASK_POS_ADM_AJUSTES, session("DIVISION_PUERTO"))

g_Najuste     = GF_PARAMETROS7("nAjuste","",6)
call addParam("nAjuste", g_Najuste, params)

g_Origen      = GF_PARAMETROS7("origen","",6)
call addParam("origen", g_Origen, params)
g_cdProducto  = GF_PARAMETROS7("cdProducto",0,6)
call addParam("cdProducto", g_cdProducto, params)
g_kilos		  = GF_PARAMETROS7("kilos","",6)
call addParam("kilos", g_kilos, params)
g_Estado	  = GF_PARAMETROS7("cmbEstado",0,6)
call addParam("cmbEstado", g_Estado, params)

fecAjusteD = GF_PARAMETROS7("fecAjusteD", "", 6)
call addParam("fecAjusteD", fecAjusteD, params)
fecAjusteM = GF_PARAMETROS7("fecAjusteM", "", 6)
call addParam("fecAjusteM", fecAjusteM, params)
fecAjusteA = GF_PARAMETROS7("fecAjusteA", "", 6)
call addParam("fecAjusteA", fecAjusteA, params)

fecAjusteDH = GF_PARAMETROS7("fecAjusteDH", "", 6)
call addParam("fecAjusteDH", fecAjusteDH, params)
fecAjusteMH = GF_PARAMETROS7("fecAjusteMH", "", 6)
call addParam("fecAjusteMH", fecAjusteMH, params)
fecAjusteAH = GF_PARAMETROS7("fecAjusteAH", "", 6)
call addParam("fecAjusteAH", fecAjusteAH, params)

concepto = GF_PARAMETROS7("concepto", "", 6)
sortBy = GF_PARAMETROS7("sortBy", "", 6)
if(sortBy = "") then sortBy = " ORDER BY A.IDAJUSTE DESC"								
call addParam("sortBy", sortBy, params)

fechaDesde = ""
if (GF_CONTROL_FECHA(fecAjusteD, fecAjusteM, fecAjusteA)) then
    fechaDesde = fecAjusteA
    fechaDesde = fechaDesde & fecAjusteM
    fechaDesde = fechaDesde & fecAjusteD
else
    fecAjusteA = ""
    fecAjusteM = ""
    fecAjusteD = ""
end if

fechaHasta = ""
if (GF_CONTROL_FECHA(fecAjusteDH, fecAjusteMH, fecAjusteAH)) then
	fechaHasta = fecAjusteAH
	fechaHasta = fechaHasta & fecAjusteMH
	fechaHasta = fechaHasta & fecAjusteDH
else
    fecAjusteAH = ""
    fecAjusteMH = ""
    fecAjusteDH = ""
end if


GP_ConfigurarMomentos
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 30
if (paginaActual = 0) then paginaActual = 1

Set rs = getAjustesPto(g_strPuerto, g_Najuste, g_Origen, g_cdProducto, g_kilos, fechaDesde, fechaHasta, g_Estado, sortBy, concepto)

Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount
%>
<HTML>
<HEAD>
	<meta http-equiv="X-UA-Compatible" content="IE=Edge">
	
	<TITLE>Poseidon - Administracion de Ajustes </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" 	type="text/css" />
	<link href="../css/main.css" 	  rel="stylesheet"	type="text/css">
	
	<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
	<script type="text/javascript" src="../scripts/paginar.js"></script>
	<script type="text/javascript" src="../scripts/controles.js"></script>
	<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	<script language="javascript">
		function bodyOnload(){		
		<% 	if (not rs.eof) then %>
				var pgn = new Paginacion("paginacion");
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "AdministracionAjustes.asp<% =params %>");						 
		<%	end if 	%>	    	
			pngfix();	
		}
		function submitInfo(){
			document.getElementById("form1").submit();
		}
		function setSortBy(pTxt){
			document.getElementById("sortBy").value = pTxt;
			submitInfo();
		}
		function asignarProducto(me){
			document.getElementById("cdProducto").value = me.value;
		}
		function asignarEstado(me){
			document.getElementById("estado").value = me.value;
		}
		function abrirCartaAjuste(IdAjuste, cdAjuste, fechaDesde, fechaHasta){
			window.open("AjusteAutorizacionPrint.asp?idAjuste=" + IdAjuste + "&cdAjuste=" + cdAjuste + "&pto=<%=g_strPuerto%>");
		}
	</script>
</HEAD>
<body onload="bodyOnload()">	

<form name="form1" id="form1">
	<div class="tableaside size100"> <!-- BUSCAR -->
		<h3> filtro - Administrar Ajustes - <% =g_strPuerto %></h3>
		
		<div id="searchfilter" class="tableasidecontent">
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("N° Ajuste") %> </div>
			<div class="col16"> <input type="text" SIZE="3" MAXLENGTH="5" id="nAjuste" name="nAjuste" value="<% =g_Najuste %>"> </div>
				 
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Kilos") %> </div>
			<div class="col16"> <input type="text" id="kilos" name="kilos" value="<% =g_Kilos %>" onkeypress="return controlIngreso(this, event, 'N')"> </div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %> </div>
			<div class="col16">
				<select id="cmbCdProducto" name="cmbCdProducto" onchange="javascript:asignarProducto(this);">
					<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
					<%
					strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM PRODUCTOS ORDER BY DSPRODUCTO"
					call GF_BD_Puertos (g_strPuerto, rsProductos, "OPEN",strSQL)										
					while not rsProductos.eof 
						if cint(g_cdProducto) = cint(rsProductos("CDPRODUCTO")) then
							mySelected = "SELECTED"
						else
							mySelected = ""
						end if	%>
							<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
					<%	rsProductos.movenext
					wend
					%>							
				</select>
				<input type="hidden" id="cdProducto" name="cdProducto" value="<%=g_cdProducto%>">
			</div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Desde") %> </div>
			<div class="col16"> 
				<input type="text" size="1" maxLength="2" value="<% =fecAjusteD %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteD" id="fecAjusteD"> /
				<input type="text" size="1" maxLength="2" value="<% =fecAjusteM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteM" id="fecAjusteM"> /
				<input type="text" size="2" maxLength="4" value="<% =fecAjusteA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteA" id="fecAjusteA">
			</div>
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Hasta") %> </div>
			<div class="col16"> 
				<input type="text" size="1" maxLength="2" value="<% =fecAjusteDH%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteDH" id="fecAjusteDH"> /
				<input type="text" size="1" maxLength="2" value="<% =fecAjusteMH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteMH" id="fecAjusteMH"> /
				<input type="text" size="2" maxLength="4" value="<% =fecAjusteAH %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAjusteAH" id="fecAjusteAH">			
			</div>	
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Estado") %> </div>
			<div class="col16"> 
				<select id="cmbEstado" name="cmbEstado" onchange="javascript:asignarEstado(this);">
					<option value="0"><%= GF_TRADUCIR("- Todos -")%></option>										
					<option value="<%=AJUSTE_ESTADO_NOAUTORIZADO %>" <%if Cint(g_Estado) = AJUSTE_ESTADO_NOAUTORIZADO then %> selected <%end if%>><%= GF_TRADUCIR("No Autorizado")%></option>
					<option value="<%=AJUSTE_ESTADO_AUTORIZADO %>" <%if Cint(g_Estado) = AJUSTE_ESTADO_AUTORIZADO then %> selected <%end if%>><%= GF_TRADUCIR("Autorizado")%></option>
					<option value="<%=AJUSTE_ESTADO_CANCELADO %>" <%if Cint(g_Estado) = AJUSTE_ESTADO_CANCELADO then %> selected <%end if%>><%= GF_TRADUCIR("Cancelado")%></option>
				</select>
				<input type="hidden" id="estado" name="estado" value="<%=g_Estado%>">	
			</div>	
			
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Concepto") %> </div>
			<div class="col16"> 
				<select id="concepto" name="concepto">
					<option value=""><%=GF_TRADUCIR("- Seleccione -") %></option>
					<option value="<%=AJUSTE_DRAFT_SURVEY %>" <% if (concepto = AJUSTE_DRAFT_SURVEY) then %> selected <% end if %> ><%=getDsCodigoAjustePuerto(AJUSTE_DRAFT_SURVEY) &" ("&AJUSTE_DRAFT_SURVEY&")" %></option>
					<option value="<%=AJUSTE_CALIDAD %>" <% if (concepto = AJUSTE_CALIDAD) then %> selected <% end if %>><%=getDsCodigoAjustePuerto(AJUSTE_CALIDAD) &" ("&AJUSTE_CALIDAD&")" %></option>
					<option value="<%=AJUSTE_MANIPULEO %>" <% if (concepto = AJUSTE_MANIPULEO) then %> selected <% end if %>><%=getDsCodigoAjustePuerto(AJUSTE_MANIPULEO) &" ("&AJUSTE_MANIPULEO&")" %></option>
					<option value="<%=AJUSTE_MERMA_VOLATIL %>" <% if (concepto = AJUSTE_MERMA_VOLATIL) then %> selected <% end if %>><%=getDsCodigoAjustePuerto(AJUSTE_MERMA_VOLATIL) &" ("&AJUSTE_MERMA_VOLATIL&")" %></option>
				</select>
			</div>	
			
			<span class="btnaction"><input type="submit" value="Buscar"></span>
		</div>
	
	</div><!-- END BUSCAR -->        
	<input type="hidden" name="sortBy" id="sortBy">
	<input type="hidden" name="sqlTipoOrder" id="sqlTipoOrder">
            
<table border="0" cellpadding="0" cellspacing="0" width="95%" align="center">
  <tr>
	  <td>
        
		</td>
	</tr>
	<tr><td><br></br></td></tr>
<% 	if (not rs.eof) then %>
	<tr>
       <td>    		  	   
	  	   <TABLE class="datagrid" align="center" width="100%">	  			
			<thead>
				<tr>					
					<th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.IDAJUSTE DESC')"><%=GF_Traducir("N° Ajuste")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.IDAJUSTE ASC')"></th>
				    <th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.CDAJUSTE DESC')"><%=GF_Traducir("Concepto")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.CDAJUSTE ASC')"></th>
				    <th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY C.DSPRODUCTO DESC')"><%=GF_Traducir("Producto")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY C.DSPRODUCTO ASC')"></th>
				    <th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.KILOSAJUSTE DESC')"><%=GF_Traducir("Kilos")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.KILOSAJUSTE ASC')"></th>
					<th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.FECHADESDE DESC')"><%=GF_Traducir("Fecha Desde")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.FECHADESDE ASC')"></th>
				    <th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.FECHAHASTA DESC')"><%=GF_Traducir("Fecha Hasta")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.FECHAHASTA ASC')"></th> 
				    <th nowrap class="reg_header_nav" align="center"><img src="../images/arrow_down.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.ESTADO DESC')"><%=GF_Traducir("Estado")%><img src="../images/arrow_up.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY A.ESTADO ASC')"></th>
				    <th nowrap class="reg_header_nav" align="center">.</th>
				</tr>
			</thead>
			<tbody>
		<%		while not rs.EOF and (reg < mostrar) 	 	
					reg = reg + 1	%>
					<tr>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="center" ><% =rs("IDAJUSTE") %></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="left" ><% =getDsCodigoAjustePuerto(rs("CDAJUSTE")) %></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="left"><% = rs("DSPRODUCTO") %></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" <% if (Cdbl(rs("KILOSAJUSTE")) < 0 ) then Response.Write " class='reg_header_rejected' "  %>align=right><% =GF_EDIT_DECIMALS(Cdbl(rs("KILOSAJUSTE")),0) & " Kg" %></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="center"><% =GF_FN2DTE(rs("FECHADESDE")) %></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="center"><% =GF_FN2DTE(rs("FECHAHASTA"))%></td>
						<td onclick="javascript:abrirCartaAjuste(<% =rs("IDAJUSTE") %>,'<% =rs("CDAJUSTE") %>','<%=rs("FECHADESDE") %>','<%=rs("FECHAHASTA") %>');" align="center"><% =getDsEstadoAjuste(rs("ESTADO")) %></td>
						<td  align="center">
						<% if (not isnull(rs("NAMEFILE"))) then %>
							<a href='../Documentos/Draft Survey/<%=g_strPuerto%>/<% = rs("NAMEFILE") %>.<% = rs("EXTFILE") %>' target='_blank'>							
								<img title='Descargar adjunto' src='../images/compras/download.png'>
							</a>
						<% end if %>	
						 </td>
					</tr>
<%                  rs.movenext
		        wend %>
			</tbody>			
			<tfoot>
				<tr>	                
					<td colspan="8"><div id="paginacion"></div></td>
				</tr>
			</tfoot>
	       </table> 
	    </td>   
	</tr>    
<%	else %>
	<tr>
		<td class="reg_Header_nav" align="center">
			<%=GF_Traducir("No se encontraron resultados")%>
		</td>
	</tr>			        
<%	end if %>
</table>
<input type="hidden" name="Pto" id="Pto" value="<%=g_strPuerto%>">
</form>
</BODY>
</HTML>