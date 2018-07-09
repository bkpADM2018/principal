<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosunificador.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<% 
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************

Dim tipo,rs,mostrar,paginaActual,totalRegistros,tipoDocumento

tipoDocumento = GF_PARAMETROS7("tipoDocumento","",6)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

    

Set sp_ret = executeSP(rs, "MERFL.MER301F1_GET_CBTES_A_FIRMAR", tipoDocumento &"||"&session("Usuario")&"||"&paginaActual&"||"&mostrar&"$$totalRegistros")

totalRegistros = sp_ret("totalRegistros")


%>
<html>
<head>
<title>Sistema de Facturación - Autorización de Facturas de Proveedores</title>
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
	var arr = Array();
	var ch= new channel();
	function bodyOnload(){
		var	tb = new Toolbar('toolbar');
		tb.addButton("toolbar-refresh","<%=GF_Traducir("Recargar")%>", "submitInfo()");
		tb.draw();
		var index = 1;
	    <%if (not rs.eof) then    
            while (not rs.eof) %> 
	            loadDocument('<% =rs("ID") %>','<% =rs("FECHA") %>','<% =rs("TIPOFORMULARIO") %>','<% =rs("EVENTO") %>',index);
	            if (ifrmSel == "") showDocument('<% =rs("ID") %>','<% =rs("FECHA") %>','<% =rs("TIPOFORMULARIO") %>','<% =rs("EVENTO") %>',index);
	            index++;
	          <%rs.MoveNext()
	        wend            
	        rs.MoveFirst()%>
            var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "submitInfo()");        
<%      end if      %>		
}	    
	function loadDocument(minuta, fecha, tipoForm, evento, index) {
        var ifrm = document.createElement("iframe");		
        ifrm.id = 'ifrm' + index;
        ifrm.name = 'ifrm' + index;
        <% if (tipoDocumento = AUTH_TYPE_PICC) then %>
            ifrm.src = 'comprasPICCPrint.asp?minuta=' + minuta + '&fecha='+fecha+'&tipoFormulario='+tipoForm+'&evento='+evento;
	    <% else %>
	        ifrm.src = 'comprasPICPrint.asp?idCotizacionElegida=' + minuta;
        <% end if %>
        document.getElementById("ifrmDiv").appendChild(ifrm);
        arr.push(index);
        hideDocument(index);
    }
	function showDocument(minuta, fecha, tipoForm, evento,index){
        document.getElementById('ifrm' + index).style.width = "100%";
        document.getElementById('ifrm' + index).style.height = "600px";		
        document.getElementById('ifrm' + index).style.display = "block";
        document.getElementById("focoMinuta").value = minuta;
        document.getElementById("focoFecha").value = fecha;
        document.getElementById("focoTipoFormulario").value = tipoForm;
        document.getElementById("focoEvento").value = evento;
        if (ifrmSel != index) {
		    hideDocument(ifrmSel);
            ifrmSel = index;
            var hkey = new Hkey('hk' + index, 'comprasFirmaPICC.asp?minuta='+minuta+'&evento='+evento+'&fecha='+fecha+'&tipoCbte='+tipoForm, '<% =HKEY() %>', 'firmar_callback()', true);
	        hkey.start();
	        if ( evento == '<%= AUTH_TYPE_PICC %>') loadInfoTipoCambio(minuta, fecha);
	        if ( evento == '<%= AUTH_TYPE_PICF %>') loadInfoFecha(minuta);
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
		    document.getElementById('hk' + ifrmSel).innerHTML =	resp + "<br>";			
		    showDocument(document.getElementById("focoMinuta").value,document.getElementById("focoFecha").value,document.getElementById("focoTipoFormulario").value,document.getElementById("focoEvento").value);
		} else {
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
	function loadInfoTipoCambio(minuta, fechaMinuta, evento){
	    ch.bind("comprasPICCInfoAjax.asp?minuta=" +  minuta +"&fecha="+fechaMinuta +"&tipoDocumento="+evento ,"CallBack_verTipoCambio()");
	    ch.send();
	}
	function CallBack_verTipoCambio(){
	    var respuesta = ch.response();
	    document.getElementById("infoTipoCambio").innerHTML = respuesta;
	}
	function loadInfoFecha(idCotizacion){
	    ch.bind("comprasPICFInfoAjax.asp?idCotizacion=" +  idCotizacion ,"CallBack_loadInfoFecha()");
	    ch.send();
	}
	function CallBack_loadInfoFecha(){
	    var respuesta = ch.response();
	    document.getElementById("infoTipoCambio").innerHTML = respuesta;
	}
</script>

    <body onload="bodyOnload()">
    <div id="toolbar"></div><br>
	    <form method="post" name="frmSel" id="frmSel" action="comprasPICCFirma.asp">	
	        <input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		    <input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>	
            <input type="hidden" name="tipoDocumento" id="tipoDocumento" value=<%=tipoDocumento%>>
	    </form>	
		<div class="tableaside size50">			
            <div><h3>COMPROBANTES A AUTORIZAR</h3><hr /></div>
			<table id="tableDoc" class="datagrid" width="80%" align="center">					
				<thead>					
					<tr>
					    <th class="thiconac" align="center" width="20%" nowrap>	<% if (tipoDocumento = AUTH_TYPE_PICC) then Response.Write "Minuta" else Response.Write "Cotizacion" %></th>
						<th class="thiconac" align="center" width="20%" nowrap>	<% =GF_TRADUCIR("Fecha") %></th>						
						<th class="thiconac" align="center" width="30%" nowrap>	<% =GF_TRADUCIR("Solicitante") %></th>
						<th class="thiconac" align="center" width="20%" nowrap><% =GF_TRADUCIR("Fecha solicitud") %></th>
						<th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Firmar") %></th>
                        <th class="thiconac" align="center" width="5%" nowrap><% =GF_TRADUCIR("Ver") %></th>
					</tr>
				</thead>	
				<tbody id="tbody"> 	
				<% if (not rs.eof) then
                        index = 1
						while (not rs.eof) %>			
						<tr>
							<td align="center"><% =rs("ID") %></td>
							<td align="center"><% =GF_FN2DTE(Trim(rs("FECHA"))) %></td>
                            <td align="center"><% =getUserDescription(rs("CDSOLICITANTE")) %></td>
                            <td align="center"><% =GF_FN2DTE(Left(rs("MMTOSOLICITUD"),8)) %></td>
							<td align="center"><div id="hk<%=index%>"></div></td>
                            <td align="center"><img src="images/search-16.png" style="cursor:pointer" onclick="javascript:showDocument('<%=rs("ID")%>','<%=rs("FECHA")%>','<%=rs("TIPOFORMULARIO")%>','<%=rs("EVENTO")%>',<%=index %>)" /></td>
						</tr>
						<%	index = index + 1
							rs.MoveNext()
						wend%>
                        <tr id="trError" style="display:none;"><td colspan="6"><div id="msjError"></div></td></tr>
                    <%
					else %>					
						<tr>
							<td colspan="6" align="center"><%=GF_TRADUCIR("No tiene minutas para autorizar")%></td>
						</tr>
				<%	end if	%>					
				</tbody>					
				
				<tfoot>
					<tr>
						<td colspan="4"><div id="paginacion"></div></td>						
					</tr>					
				</tfoot>				
			</table>
            <input id="focoMinuta" name="focoMinuta" type="hidden"/>
            <input id="focoFecha" name="focoFecha" type="hidden"/>
            <input id="focoTipoFormulario" name="focoTipoFormulario" type="hidden"/>
            <input id="focoEvento" name="focoEvento" type="hidden"/>
            <div> 
			    <iframe id='ifrmContable' width="100%" style="border: none;" src='' onload="autoAjustIframe(this)"></iframe>
			</div>
            <div id='infoTipoCambio' width="100%"></div>
		</div>	
		
		<div id="ifrmDiv" class="tableaside size50"> </div>		
    </body>        
</html>