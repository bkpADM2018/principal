<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosTitulos.asp"-->
<!--#include file="../../includes/procedimientosSQL.asp"-->
<%


'----------------------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'----------------------------------------------------------------------------------------------------------------------
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
'----------------------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto, Conn, params,lineasTotales, mostrar,paginaActual,g_cdProducto,flagPermiso

g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)
flagPermiso = true
if (leerPermisos(g_strPuerto, TASK_PRODUCT_USER) = NO_TIENE_PERMISO) then flagPermiso = false

g_cdProducto = GF_PARAMETROS7("cmbCdProducto",0,6)
call addParam("cmbCdProducto", g_cdProducto, params)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

strSQL = "Select * from dbo.productos  "
if (g_cdProducto <> 0) then strSQL = strSQL & "where cdproducto = " & g_cdProducto
strSQL = strSQL & " order by cdproducto "
call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)	

Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount
%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administracion de Productos </TITLE>
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>
</HEAD>
<script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../../scripts/paginar.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script language="javascript">
	var ch = new channel();
	var up1;	
	function onLoadPage(){
		tb = new Toolbar('toolbar', 6,'../../images/');				
		tb.addButton("refresh-16.png", "Refrescar", "submitInfo('<%=ACCION_SUBMITIR%>')");
		<% if(flagPermiso) then %>
        tb.addButton("add-16.png", "Nuevo Producto", "nuevoProducto()");
        <% end if %>
		tb.draw();
		<% 	if (not rs.eof) then %>
			var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "productoAdministrar.asp<% =params %>");
		<%	end if 	%>	   
		
	}	
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();
	}
	function nuevoProducto(){				
		document.location.href = "productoAgregar.asp?pto=<%=g_strPuerto%>";
	}
	function editarProducto(pProducto){		
		document.location.href = "productoAgregar.asp?pto=<%=g_strPuerto%>&cdProducto="+pProducto;
	}
	function eliminarProducto(pCdProducto){
		if (confirm("Desea eliminar el Producto")){
			ch.bind("productoAjax.asp?pto=<%=g_strPuerto %>&cdProducto="+pCdProducto+"&accion=<%=ACCION_BORRAR%>", "eliminarProducto_Callback("+ pCdProducto +")");
			ch.send();
		}
	}
	function eliminarProducto_Callback(pCdProducto){
		submitInfo('<%=ACCION_SUBMITIR%>');
	}
	
</script>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post>
<div class="tableaside size100"> <!-- BUSCAR -->
<h3> Filtro de Productos</h3>        
<div id="searchfilter" class="tableasidecontent">
	<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %> </div>
        <div class="col16"> 
			<select id="cmbCdProducto" name="cmbCdProducto" >
										<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
										<%
										strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
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
		</div>						
			<span class="btnaction"><input type="button" value="Buscar" id=cmdSearch name=cmdSearch onclick="submitInfo('<%=ACCION_SUBMITIR%>');"></span>
    </div>
</div><!-- END BUSCAR -->
					
<div class="col66"></div>

		
								
	<table width="100%" class="datagrid" align="center" id="tblResult" name="tblResult">           
<% 	if (not rs.eof) then %>	
			<thead>
                <tr>
					<th width="10%" align="left"><% =GF_TRADUCIR("Codigo") %></th>	
                    <th width="65%" align="left"><% =GF_TRADUCIR("Descripcion") %></th>	
				    <th width="15%"  align="left"><% =GF_TRADUCIR("Descripcion Abr.") %></th>
				    <th width="5%"  align="center">.</th>
				    <th width="5%"  align="center">.</th>
                </tr>
            </thead> 
				<tbody>
		<%		while not rs.EOF and (reg < mostrar)
					reg = reg + 1	    %>
					<tr class="reg_Header_navdos">		
						<td align="left"><font size="2">
							<% =rs("CDPRODUCTO")%></font>
						</td>	
						<td align="left"><font size="2">
							<% =rs("DSPRODUCTO")%></font>
						</td>	
						<td align="left"><font size="2">
							<% =rs("DSPRODUCTOABR") %></font>
						</td>
                        <% if (flagPermiso) then %>
						<td align="center"><img src="../../images/edit-16.png" title="Editar" id="editarProd" style="cursor:pointer;" onclick="javascript:editarProducto(<%=rs("CDPRODUCTO")%>)"></td>						
						<td align="center"><img src="../../images/cross-16.png" title="Eliminar" id="eliminarProd" style="cursor:pointer;" onclick="eliminarProducto(<%=rs("CDPRODUCTO")%>)"></td>						
                        <% else %>
                        <td align="center"><img src="../../images/see-16.png" title="Ver" style="cursor:pointer;" onclick="javascript:editarProducto(<%=rs("CDPRODUCTO")%>)"></td>						
						<td align="center"></td>
                        <% end if %>
					</tr>
					<tr class="troculto" id="TR_ID_<% =rs("CDPRODUCTO") %>">
						<td colspan="3">
							<div align="center" id="TBL_<% =rs("CDPRODUCTO") %>" style="position:absolute; "><img src='../../images/Loading4.gif'></div>
						</td>
					</tr>
<%                  rs.movenext()
		        wend %>
			</tbody>			
            				
            <tfoot>
                <td colspan="5"><div id="paginacion"></div></td>
            </tfoot>
<%	else %>
		<tr><td align="center"><%=GF_TRADUCIR("No se encontraron productos")%></td></tr>
			
<%	end if %>
</TABLE>
<input type="hidden" name="accion" id="accion" value="<%= accion %>">
</form>
</BODY>
</HTML>