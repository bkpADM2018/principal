<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosTitulos.asp"-->
<!--#include file="../includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<%
Function getMermaVolatil(p_fecha ,p_cdProducto, p_cdCliente, p_cdSilo,p_Pto)
    Dim strSQL
    strSQL = "SELECT  (YEAR(MV.DTCONTABLE)*10000 + Month(MV.DTCONTABLE)*100 + DAY(MV.DTCONTABLE)) AS FECHA, "&_
			 "	       MV.CDPRODUCTO, "&_
			 "	       PRO.DSPRODUCTO, "&_
			 "	       MV.CDCLIENTE, "&_
			 "	       CLI.DSCLIENTE, "&_
			 "	       MV.CDSILO, "&_
			 "	       MV.RATIO, "&_
			 "	       MV.CDUSUARIO, "&_
			 "	       MV.MMTO "&_
		     "FROM DBO.TBLREGLASMERMAVOLATIL MV "&_
			 "   LEFT JOIN PRODUCTOS PRO ON MV.CDPRODUCTO = PRO.CDPRODUCTO "&_
			 "   LEFT JOIN CLIENTES CLI ON MV.CDCLIENTE = CLI.CDCLIENTE "&_
		     "WHERE 1 = 1 "
			 if (p_fecha <> "") then strSQL = strSQL & " AND MV.DTCONTABLE = '"& p_fecha &"' "
             if (Cdbl(p_cdProducto) <> 0) then strSQL = strSQL & " AND MV.CDPRODUCTO = "& p_cdProducto 
             if (Cdbl(p_cdCliente) <> 0) then strSQL = strSQL & " AND MV.CDCLIENTE = "& p_cdCliente 
	         if (p_cdSilo <> "") then strSQL = strSQL & " AND MV.CDSILO = '"& p_cdSilo &"' "
             strSQL = strSQL & " ORDER BY DTCONTABLE DESC,CDPRODUCTO,CDCLIENTE,CDSILO "
    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL) 
    Set getMermaVolatil = rs
End Function
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

Call initTaskAccessInfo(TASK_POS_MERMA_VOLATIL, session("DIVISION_PUERTO"))

g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)

g_cdProducto = GF_PARAMETROS7("cmbCdProducto",0,6)
call addParam("cmbCdProducto", g_cdProducto, params)

g_fechaD = GF_PARAMETROS7("fechaD", 0, 6)
call addParam("fechaD", g_fechaD, params)
g_fechaM = GF_PARAMETROS7("fechaM", 0, 6)
call addParam("fechaM", g_fechaM, params)
g_fechaA = GF_PARAMETROS7("fechaA", 0, 6)
call addParam("fechaA", g_fechaA, params)
g_fecha = ""
if ((g_fechaA <> 0)and(g_fechaM <> 0)and(g_fechaD <> 0)) then 
    g_fecha = g_fechaA &"-"& g_fechaM &"-"& g_fechaD
end if

g_cdCliente = GF_PARAMETROS7("cdCliente", 0, 6)
call addParam("cdCliente", g_cdCliente, params)
g_dsCliente = GF_PARAMETROS7("dsCliente", "", 6)
call addParam("dsCliente", g_dsCliente, params)

g_cdSilo = GF_PARAMETROS7("cdSilo", "", 6)
call addParam("cdSilo", g_cdSilo, params)

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (Cint(mostrar) = 0) then mostrar = 10
if (Cint(paginaActual) = 0) then paginaActual = 1


Set rs = getMermaVolatil(g_fecha ,g_cdProducto, g_cdCliente, g_cdSilo, g_strPuerto)
Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount

%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administracion de Merma Volatil </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../css/main.css" type="text/css">	
    <link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">	
	<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>
<script type="text/javascript" src="../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script language="javascript">
	var ch = new channel();
	var up1;	
	function onLoadPage(){
		tb = new Toolbar('toolbar', 6,'../images/');				
		tb.addButton("refresh-16.png", "Refrescar", "submitInfo('<%=ACCION_SUBMITIR%>')");
		tb.addButton("add-16.png", "Agregar Merma", "nuevaMermaVolatil()");
		tb.draw();
        <% 	if (not rs.eof) then %>
			var pgn = new Paginacion("paginacion");
            pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "mermaVolatilAdministrar.asp<% =params %>");
        <%	end if 	%>
        autoCompleteCorredor();
		
	}	
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();
	}
	function eliminarMermaVolatil(pFecha,pCdProducto,pCdCliente,pCdSilo){
	    if (confirm("Desea eliminar la merma volatil?")){
	        ch.bind("mermaVolatilAjax.asp?Pto=<%=g_strPuerto %>&fecha="+pFecha+"&cdProducto="+pCdProducto+"&cdCliente="+pCdCliente+"&cdSilo="+pCdSilo+"&accion=<%=ACCION_BORRAR%>", "eliminarMermaVolatil_Callback()");
			ch.send();
		}
	}
	function eliminarMermaVolatil_Callback(){
		submitInfo('<%=ACCION_SUBMITIR%>');
	}
	function CerrarCal(cal) {
	    cal.hide();
	}		
	function MostrarCalendario(p_objID, funcSel) {
	    var dte= new Date();		    	    
	    var elem= document.getElementById(p_objID);
	    if (calendar != null) calendar.hide();		
	    var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
	    cal.setRange(1993, 2045);
	    cal.create();
	    calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	function SeleccionarCal(cal, date) {
	    var str= new String(date);		
	    document.getElementById("dtFecha").value = str;
	    document.getElementById("fechaD").value = str.substr(0,2);
	    document.getElementById("fechaM").value = str.substr(3,2);
	    document.getElementById("fechaA").value = str.substr(6,4);
	    if (cal) cal.hide();
	}	

	function autoCompleteCorredor(){
	    $( "#dsCliente" ).autocomplete({
	        minLength: 2,
	        source: "puertosStreamElementos.asp?tipo=JQClientes&pto=<%=g_strPuerto%>",
	        focus: function( event, ui ) {
	            $( "#dsCliente").val(ui.item.dscliente);
	            return false;
	        },
	        select: function( event, ui ) {
	            $( "#dsCliente"    ).val (ui.item.dscliente);
	            $( "#cdCliente"    ).val (ui.item.cdcliente);
	            return false;
	        },
	        change: function( event, ui ) {				
	            if (!ui.item) {					
	                $( "#dsCliente").val ("");
	                $( "#cdCliente").val ("");
	            }
	        }
	    })
		.data( "autocomplete" )._renderItem = function( ul, item ) {
		    return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdcliente + " - <font style='font-size:10;'>" + item.dscliente + "</font></a>" )
				.appendTo( ul );
		};
	}		
	function nuevaMermaVolatil(){
	    myPopUp = new winPopUp('Iframe', 'mermaVolatilPopUp.asp?pto=<%=g_strPuerto %>', '700', '300', 'Nuevo',"submitInfo('<%=ACCION_SUBMITIR%>')");
	}
	function editarMermaVolatil(pFecha,pCdProducto,pCdCliente,pCdSilo){
	    myPopUp = new winPopUp('Iframe', "mermaVolatilPopUp.asp?pto=<%=g_strPuerto %>&fecha="+pFecha+"&cdProducto="+pCdProducto+"&cdCliente="+pCdCliente+"&cdSilo="+pCdSilo, '700', '300', 'Editar',"submitInfo('<%=ACCION_SUBMITIR%>')");
	}
</script>
</HEAD>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post action="mermaVolatilAdministrar.asp">
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> Filtros </h3>        
    <div id="searchfilter" class="tableasidecontent">
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Valido desde") %> </div>
        <div class="col16"> 
            <table>
				<tr><td>
				    <input type="text" name="dtFecha" id="dtFecha" readonly onclick="javascript:MostrarCalendario('dtFecha', SeleccionarCal)" value="<% =g_fecha %>">
				</td></tr>
				<input type="hidden" id="fechaD" name="fechaD" value="<%=g_fechaD %>">
				<input type="hidden" id="fechaM" name="fechaM" value="<%=g_fechaM %>">
				<input type="hidden" id="fechaA" name="fechaA" value="<%=g_fechaA %>">
			</table>
        </div>
	    <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %> </div>
        <div class="col16">
	        <select id="cmbCdProducto" name="cmbCdProducto" >
		        <option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
			    <%  strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
			        call GF_BD_Puertos (g_strPuerto, rsProductos, "OPEN",strSQL)
				    while not rsProductos.eof
				        if cint(g_cdProducto) = cint(rsProductos("CDPRODUCTO")) then
					        mySelected = "SELECTED"
					    else
					        mySelected = ""
					    end if	%>
					    <option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
				    <%	rsProductos.movenext
				    wend %>
            </select>									
	    </div>	
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Cliente") %> </div>
        <div class="col16">
            <input type="text"   name="dsCliente" id="dsCliente" value="<%=g_dsCliente %>" style="width:150px">
			<input type="hidden" name="cdCliente" id="cdCliente" value="<%=g_cdCliente %>">
        </div>
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Silo") %> </div>
        <div class="col16">
            <input type="text" name="cdSilo" id="cdSilo" value="<%=g_cdSilo %>" style="width:150px">
        </div>
	    <span class="btnaction"><input type="button" value="Buscar" id=cmdSearch name=cmdSearch onclick="submitInfo('<%=ACCION_SUBMITIR%>');"></span>
    </div>
</div><!-- END BUSCAR -->
					
<div class="col66"></div>
								
	<table width="85%" class="datagrid" align="center" id="tblResult" name="tblResult">           
        <thead>
            <tr>
			    <th width="15%" align="center"><% =GF_TRADUCIR("Valido desde") %></th>
                <th width="25%" align="center"><% =GF_TRADUCIR("Producto") %></th>
				<th width="25%" align="center"><% =GF_TRADUCIR("Cliente") %></th>
                <th width="15%" align="center"><% =GF_TRADUCIR("Silo") %></th>
				<th width="10%" align="center"><% =GF_TRADUCIR("Ratio") %></th>
				<th width="5%"  align="center">.</th>
                <th width="5%"  align="center">.</th>
            </tr>
        </thead>
    <% 	if (not rs.eof) then %>	
		<tbody>
    <%	while not rs.EOF and (reg < mostrar)
		    reg = reg + 1	    %>
            <tr class="reg_Header_navdos">		
			    <td align="center"><font size="2"><% =GF_FN2DTE(rs("FECHA")) %></font></td>	
				<td align="left"><font size="2"><% =rs("DSPRODUCTO")%></font></td>	
                <td align="left"><font size="2"><% =Trim(rs("DSCLIENTE")) %></font></td>
                <td align="left"><font size="2"><% =Trim(rs("CDSILO")) %></font></td>
                <td align="rigth"><font size="2"><% =GF_EDIT_DECIMALS(Cdbl(rs("RATIO"))*100,2) & " %" %></font></td>
				<td align="center"><img src="../images/edit-16.png" title="Editar" id="editarProd" style="cursor:pointer;" onclick="javascript:editarMermaVolatil('<%=rs("FECHA")%>',<%=rs("CDPRODUCTO") %>,<%=rs("CDCLIENTE") %>,'<%=rs("CDSILO") %>')"></td>
				<td align="center"><img src="../images/cross-16.png" title="Eliminar" id="eliminarProd" style="cursor:pointer;" onclick="eliminarMermaVolatil('<%=rs("FECHA")%>',<%=rs("CDPRODUCTO") %>,<%=rs("CDCLIENTE") %>,'<%=rs("CDSILO") %>')"></td>
			</tr>
        <%  rs.movenext()
		wend %>
        </tbody>			
        <tfoot>
            <tr><td colspan="7"><div id="paginacion"></div></td></tr>
        </tfoot>
<%	else %>
		<tr><td align="center" colspan="7"><%=GF_TRADUCIR("No se encontraron mermas")%></td></tr>
<%	end if %>

</TABLE>
<input type="hidden" name="accion" id="accion" value="<%= accion %>">
<input type="hidden" name="Pto" id="Pto" value="<%= g_strPuerto %>">
<input type="hidden" id="registrosPorPagina" name="registrosPorPagina" value="<% =mostrar %>">
<input type="hidden" id="numeroPagina" name="numeroPagina" value="<% =paginaActual %>">
</form>
</BODY>
</HTML>