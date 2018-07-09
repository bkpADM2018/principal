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



Set sp_ret = executeSP(rs, "EJIFL.TBLPROVISIONESCANE_GET_CBTES_A_FIRMAR", nroLote &"||"& fechaLote &"||"& session("Usuario") &"||"& paginaActual&"||"& mostrar &"$$totalRegistros")
totalRegistros = sp_ret("totalRegistros")

%>
<html>
<head>
<title>SISTEMA DE PROVISIONES - Autorizaci&oacuten de proviciones desde cancelaci&oacuten autom&aacutetica</title>
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
    var arr = Array();
    var ch= new channel();
    function bodyOnload(){
        var	tb = new Toolbar('toolbar');
        tb.addButton("toolbar-refresh","<%=GF_Traducir("Recargar")%>", "submitInfo()");
        tb.draw();
        var index = 1;
        <%if (not rs.eof) then    
            while (not rs.eof) %> 
	            loadDocument('<% =rs("NROLOTE") %>','<% =rs("FECHALOTE") %>',index);
                if (ifrmSel == "") showDocument('<% =rs("NROLOTE") %>','<% =rs("FECHALOTE") %>',index);
                index++;
                <%rs.MoveNext()
            wend            
            rs.MoveFirst()%>
            var pgn = new Paginacion("paginacion");
            pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "submitInfo()");        
        <% end if %>		
    }

    function loadDocument(p_NroLote, p_FechaLote, index) {
        var ifrm = document.createElement("iframe");		
        ifrm.id = 'ifrm' + index;
        ifrm.name = 'ifrm' + index;
        ifrm.src = 'provisionesCancelacionAutomaticaPrint.asp?nroLote=' + p_NroLote +"&fechaLote="+p_FechaLote;
        document.getElementById("ifrmDiv").appendChild(ifrm);
        arr.push(index);
        hideDocument(index);
    }
    function showDocument(p_NroLote, p_FechaLote, index){
        document.getElementById('ifrm' + index).style.width = "100%";
        document.getElementById('ifrm' + index).style.height = "600px";		
        document.getElementById('ifrm' + index).style.display = "block";
        document.getElementById("focoNroLote").value = p_NroLote;
        document.getElementById("focoFechaLote").value = p_FechaLote;
        if (ifrmSel != index) {
            hideDocument(ifrmSel);
            ifrmSel = index;
            var hkey = new Hkey('hk' + index, 'provisionesFirma.asp?nroLote='+p_NroLote+'&fechaLote='+p_FechaLote, '<% =HKEY() %>', 'firmar_callback()', true);
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
	    <form method="post" name="frmSel" id="frmSel" action="provisionesCancelacionAutomaticaFirma.asp">	
	        <input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		    <input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>	
            <input type="hidden" name="nroLote" id="nroLote" value=<%=nroLote%>>
            <input type="hidden" name="fechaLote" id="fechaLote" value=<%=fechaLote%>>
	    </form>	
		<div class="tableaside size50">			
            <div><h3>COMPROBANTES A AUTORIZAR</h3><hr /></div>
			<table id="tableDoc" class="datagrid" width="90%" align="center">					
				<thead>					
					<tr>
					    <th class="thiconac" align="center" width="15%" nowrap>	<% =GF_TRADUCIR("Nro.Lote") %></th>
						<th class="thiconac" align="center" width="15%" nowrap>	<% =GF_TRADUCIR("Fecha Lote") %></th>						
						<th class="thiconac" align="center" width="20%" nowrap>	<% =GF_TRADUCIR("Provisiones") %></th>
						<th class="thiconac" align="center" width="20%" nowrap><% =GF_TRADUCIR("Gastos") %></th>
                        <th class="thiconac" align="center" width="20%" nowrap><% =GF_TRADUCIR("Gastos a cancelar") %></th>
						<th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Firmar") %></th>
                        <th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Ver") %></th>
					</tr>
				</thead>	
				<tbody id="tbody"> 	
				<% if (not rs.eof) then
                        index = 1
						while (not rs.eof) %>			
						<tr>
							<td align="center"><% =rs("NROLOTE") %></td>
							<td align="center"><% =GF_FN2DTE(rs("FECHALOTE")) %></td>
                            <td align="right"><% =TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(rs("TOTALPROVISIONPESOS"))*100,2) %></td>
                            <td align="right"><% =TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(rs("TOTALGASTOPESOS"))*100,2) %></td>
                            <td align="right"><% =TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(rs("TOTALCANCELACIONPESOS"))*100,2) %></td>
							<td align="center"><div id="hk<%=index%>"></div></td>
                            <td align="center"><img src="images/search-16.png" style="cursor:pointer" onclick="javascript:showDocument(<%=rs("NROLOTE")%>,<%=rs("FECHALOTE")%>,<%=index%>)" /></td>
						</tr>
						<%	index = index + 1
							rs.MoveNext()
						wend%>
                        <tr id="trError" style="display:none;"><td colspan="7"><div id="msjError"></div></td></tr>
                    <%
					else %>					
						<tr>
							<td colspan="7" align="center"><%=GF_TRADUCIR("No tiene minutas para autorizar")%></td>
						</tr>
				<%	end if	%>					
				</tbody>
				<tfoot>
                    <tr><td colspan="7">&nbsp</td></tr>
					<tr>
						<td colspan="7"><div id="paginacion"></div></td>						
					</tr>					
				</tfoot>				
			</table>
            <input id="focoNroLote" name="focoNroLote" type="hidden"/>
            <input id="focoFechaLote" name="focoFechaLote" type="hidden"/>
		</div>
		<div id="ifrmDiv" class="tableaside size50"> </div>		
    </body>        
</html>
