<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<% 
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************

Dim nroLote,fechaLote,rs,mostrar,paginaActual,totalRegistros

nroLote = GF_PARAMETROS7("nroLote",0,6)
fechaLote = GF_PARAMETROS7("fechaLote",0,6)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACION_GET_CBTES_A_FIRMAR", session("Usuario") &"||"& paginaActual&"||"& mostrar &"$$totalRegistros")
totalRegistros = sp_ret("totalRegistros")

%>
<html>
<head>
<title>SISTEMA DE COMPRAS - Autorizaci&oacuten de Partidas Presupuestarias</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript">	
	
    var ifrmSel = "";	
    var ch= new channel();
    function bodyOnload(){
        var	tb = new Toolbar('toolbar');
        tb.addButton("toolbar-refresh","<%=GF_Traducir("Recargar")%>", "submitInfo()");
        tb.draw();
        var index = 1;
        <%if (not rs.eof) then    
            while (not rs.eof) %> 
	            loadDocument('<% =rs("IDOBRA") %>',index);
                if (ifrmSel == "") showDocument('<% =rs("IDREASIGNACION") %>',index);
                index++;
                <%rs.MoveNext()
            wend            
            rs.MoveFirst()%>
            var pgn = new Paginacion("paginacion");
            pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "submitInfo()");        
        <% end if %>		
    }

    function loadDocument(p_IdObra, index) {
        var ifrm = document.createElement("iframe");		
        ifrm.id = 'ifrm' + index;
        ifrm.name = 'ifrm' + index;
        ifrm.src = 'comprasBudgetObraPrint.asp?idObra=' + p_IdObra;
        document.getElementById("ifrmDiv").appendChild(ifrm);
        hideDocument(index);
    }
    function showDocument(p_IdReasignacion, index){
        document.getElementById('ifrm' + index).style.width = "100%";
        document.getElementById('ifrm' + index).style.height = "600px";		
        document.getElementById('ifrm' + index).style.display = "block";
        document.getElementById("focoId").value = p_IdReasignacion;
        if (ifrmSel != index) {
            hideDocument(ifrmSel);
            ifrmSel = index;
            var hkey = new Hkey('hk' + index, 'comprasFirmarBudget.asp?idReasignacion='+p_IdReasignacion, '<% =HKEY() %>', 'firmar_callback()', true);
            hkey.start();
        }
		
    }
    function hideDocument(id){
        if (id != "") {
            document.getElementById('ifrm' + id).style.width = "0px";
            document.getElementById('ifrm' + id).style.height = "0px";		
            document.getElementById('ifrm' + id).style.display = "none";	
            document.getElementById('hk' + id).innerHTML = "";
        }
    }   
    function firmar_callback(resp){
        if (resp != "<%=RESPUESTA_OK%>") alert(resp);
        document.getElementById("frmSel").submit();
    }         
    	  
    
    function submitInfo(pg, lpp) {                
        document.getElementById("numeroPagina").value= pg;
        document.getElementById("registrosPorPagina").value= lpp;
        document.getElementById("frmSel").submit();
    }
      
    
    function autoAjustIframe(id){
        var valueScrollHeight = id.contentDocument.body.scrollHeight;		
        if(valueScrollHeight == 0) valueScrollHeight = 30;
        if(document.getElementById) id.style.height = valueScrollHeight+'px';
    }
    

</script>

    <body onload="bodyOnload()">
    <div id="toolbar"></div><br>
	    <form method="post" name="frmSel" id="frmSel" action="comprasBudgetFirma.asp">	
	        <input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		    <input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>	
	    </form>	
		<div class="tableaside size50">			
            <div><h3>COMPROBANTES A AUTORIZAR</h3><hr /></div>
			<table id="tableDoc" class="datagrid" width="90%" align="center">					
				<thead>					
					<tr>
					    <th class="thiconac" align="center" width="6%" nowrap>	<% =GF_TRADUCIR("Id Ajuste") %></th>
						<th class="thiconac" align="center" width="10%" nowrap>	<% =GF_TRADUCIR("Fecha") %></th>						
						<th class="thiconac" align="center" width="32%" nowrap>	<% =GF_TRADUCIR("Obra") %></th>
						<th class="thiconac" align="center" width="13%" nowrap><% =GF_TRADUCIR("Detalle Origen") %></th>
                        <th class="thiconac" align="center" width="13%" nowrap><% =GF_TRADUCIR("Detalle Destino") %></th>
                        <th class="thiconac" align="center" width="18%" nowrap><% =GF_TRADUCIR("Importe") %></th>
						<th class="thiconac" align="center" width="4%" nowrap><% =GF_TRADUCIR("Firmar") %></th>
                        <th class="thiconac" align="center" width="4%" nowrap><% =GF_TRADUCIR("Ver") %></th>
					</tr>
				</thead>	
				<tbody id="tbody"> 	
				<% if (not rs.eof) then
                        index = 1
						while (not rs.eof) %>			
						<tr>
							<td align="center"><% =rs("IDREASIGNACION") %></td>
							<td align="center"><% =GF_FN2DTE(rs("FECHA")) %></td>
                            <td align="left"><%= rs("DSOBRA") %></td>
                            <td align="center">
                                <% if (Cdbl(rs("IDAREAORIGEN")) <> 0 or Cdbl(rs("IDDETALLEORIGEN")) <> 0) then
                                      Response.Write rs("IDAREAORIGEN") &"-"& rs("IDDETALLEORIGEN") 
                                   end if %>
                            </td>
                            <td align="center"><%= rs("IDAREADESTINO") &"-"& rs("IDDETALLEDESTINO") %></td>                            
                            <td align="right"><% =TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(rs("IMPORTEDOLARES"),2) %></td>                            
							<td align="center"><div id="hk<%=index%>"></div></td>
                            <td align="center"><img src="images/search-16.png" style="cursor:pointer" onclick="javascript:showDocument(<%=rs("IDREASIGNACION")%>,<%=index%>)" /></td>
						</tr>
						<%	index = index + 1
							rs.MoveNext()
						wend%>
                        <tr id="trError" style="display:none;"><td colspan="7"><div id="msjError"></div></td></tr>
                    <%
					else %>					
						<tr>
							<td colspan="8" align="center"><%=GF_TRADUCIR("No tiene minutas para autorizar")%></td>
						</tr>
				<%	end if	%>					
				</tbody>
				<tfoot>
                    <tr><td colspan="8">&nbsp</td></tr>
					<tr>
						<td colspan="8"><div id="paginacion"></div></td>						
					</tr>					
				</tfoot>				
			</table>
            <input id="focoId" name="focoId" type="hidden"/>
		</div>
		<div id="ifrmDiv" class="tableaside size50"> </div>		
    </body>        
</html>
