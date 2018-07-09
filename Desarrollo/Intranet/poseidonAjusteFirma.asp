<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosunificador.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="interfacturas/interfacturas.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<% 
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim tipo,rs,mostrar,paginaActual,totalRegistros,pto

pto = GF_PARAMETROS7("pto","",6)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1
                                                
Call executeProcedureDb(pto, rs, "TBLAJUSTES_GET_CBTES_A_FIRMAR", Session("Usuario") &"||")
lineasTotales = rs.recordcount
Call setupPaginacion(rs, paginaActual, mostrar)

%>
<html>
<head>
<title>Sistema Poseidon - Autorización de Ajustes</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">	
	
	var ifrmSel = "";	
	var ch= new channel();
	function bodyOnload(){
		var	tb = new Toolbar('toolbar');
		tb.addButton("toolbar-refresh","<%=GF_Traducir("Recargar")%>", "submitInfo()");
		tb.draw();	
		<% 	if (not rs.eof) then %>
			
		<%	end if 	%>	    	        
<%      if (not rs.eof) then            
			index = 0
            while ((not rs.eof) and (index < mostrar)) 
                index = index + 1		%>
                loadDocument('<% =rs("ID") %>', <% =index %>);
	            if (ifrmSel == "") showDocument('<% =rs("ID") %>', '<% =rs("CDAJUSTE") %>', <% =index %>);
<%              rs.MoveNext()
            wend            
            rs.MoveFirst()
%>
            var pgn = new Paginacion("paginacion");
            pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "submitInfo()");						 
<%      end if      %>		
}	    
	function loadDocument(pIdAjuste, pIndex) {
        var ifrm = document.createElement("iframe");		
        ifrm.id = 'ifrm' + pIndex;
        ifrm.name = 'ifrm' + pIndex;
        var strParameter = '?pto=<%= pto %>&idAjuste='+ pIdAjuste;
        
        ifrm.src = 'poseidonAjusteAutorizacionPrint.asp' + strParameter;
        document.getElementById("ifrmDiv").appendChild(ifrm);
        hideDocument(pIndex);
    }
	function showDocument(pId,pCdAjuste,pIndex){
	    document.getElementById('ifrm' + pIndex).style.width = "100%";
        document.getElementById('ifrm' + pIndex).style.height = "600px";		
        document.getElementById('ifrm' + pIndex).style.display = "block";
        if (ifrmSel != pIndex) {
		    hideDocument(ifrmSel);
		    ifrmSel = pIndex;
		    var hkey = new Hkey('hk' + pIndex, 'poseidonAjusteFirmaAjax.asp?pto=<%= pto %>&idAjuste='+ pId +'&cdAjuste='+pCdAjuste, '<% =HKEY() %>', 'firmar_callback()', true);
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
		if (resp != "<%=RESPUESTA_OK%>"){	
		    document.getElementById("divError").style.display = "block";
		    document.getElementById("divError").innerHTML = resp;
		} else {
		    document.getElementById("divError").style.display = "none";
		    document.getElementById("frmSel").submit();
		}
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
    <div><h3>COMPROBANTES A AUTORIZAR</h3><hr /></div>
	    <form method="post" name="frmSel" id="frmSel" action="poseidonAjusteFirma.asp">	
	        <input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		    <input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>	
            <input type="hidden" name="pto" id="pto" value=<%=pto%>>	
	    </form>	
		<div class="tableaside size50">		
            <div style="width:80%;display:none;margin:0 auto;" class="errormsj" id="divError"></div>	
			<table id="tableDoc" class="datagrid" width="80%" align="center">					
				<thead>					
					<tr>
					    <th class="thiconac" align="center" width="30%" nowrap>	<% =GF_TRADUCIR("Tipo Ajuste") %></th>
						<th class="thiconac" align="center" width="20%" nowrap>	<% =GF_TRADUCIR("Fecha desde") %></th>
                        <th class="thiconac" align="center" width="20%" nowrap>	<% =GF_TRADUCIR("Fecha hasta") %></th>
						<th class="thiconac" align="center" width="15%" nowrap><% =GF_TRADUCIR("Firmar") %></th>
						<th class="thiconac" align="center" width="15%" nowrap><% =GF_TRADUCIR("Ver") %></th>
					</tr>
				</thead>	
				<tbody id="tbody"> 	
				<% if (not rs.eof) then
						reg = 0
                        while not rs.EOF and (reg < mostrar)
                            reg = reg + 1 %>
						    <tr>
							    <td align="left"><% =getDsCodigoAjustePuerto(rs("CDAJUSTE")) & " ("& rs("CDAJUSTE") &")" %></td>
                                <td align="center"><%=GF_FN2DTE(rs("FECHADESDE"))%></td>
                                <td align="center"><%=GF_FN2DTE(rs("FECHAHASTA"))%></td>
							    <td align="center"><div id="hk<%=reg %>"></div></td>
							    <td align="center"><img src="images/search-16.png" style="cursor:pointer" onclick="javascript:showDocument(<%=rs("ID")%>,'<% =rs("CDAJUSTE") %>',<%=reg %>)" /></td>
						    </tr>
						<%	rs.MoveNext()
						wend
					else %>					
						<tr>
							<td colspan="5" align="center"><%=GF_TRADUCIR("No tiene ajustes por firmar")%></td>
						</tr>
				<%	end if	%>					
				</tbody>					
				
				<tfoot>
					<tr>
						<td colspan="5"><div id="paginacion"></div></td>						
					</tr>					
				</tfoot>				
			</table>
		</div>	
		
		<div id="ifrmDiv" class="tableaside size50"> </div>		
    </body>        
</html>