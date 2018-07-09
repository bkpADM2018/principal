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

Dim tipo,rs,mostrar,paginaActual,totalRegistros

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

                                                
Set sp_ret = executeSP(rs, "TFFL.TF100F1_GET_CBTES_A_FIRMAR", session("Usuario")&"||"&paginaActual&"||"&mostrar&"$$totalRegistros")

totalRegistros = sp_ret("totalRegistros")
%>
<html>
<head>
<title>Sistema de Facturación - Autorización de Facturas</title>
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
<%      if (not rs.eof) then            
            while (not rs.eof)             %>                
	            loadDocument('<% =rs("FCRGNR") %>');
	            if (ifrmSel == "") showDocument('<% =rs("FCRGNR") %>');		        
<%              
                rs.MoveNext()
            wend            
            rs.MoveFirst()            
%>
            var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "submitInfo()");        
			stratSignAll('<% =FAC_CODIGO_CONCEPTO_GP %>');
<%      end if      %>		
}	    
    function loadDocument(id) {
        var ifrm = document.createElement("iframe");		
        ifrm.id = 'ifrm' + id;
        ifrm.name = 'ifrm' + id;         
        ifrm.src = 'interfacturas/interfacturasPrintCommon.asp?lote=' + id + '&tipoPDF=<%=PDF_STREAM_MODE%>';
        document.getElementById("ifrmDiv").appendChild(ifrm);
        arr.push(id);
        hideDocument(id);
    }
    function stratSignAll(cd) {
        var hkey = new Hkey('hkAll' + cd, 'interfacturas/interfacturasFirmaAllAjax.asp?cd=' + cd , '<% =HKEY() %>', "firmarAll_callback('" + cd + "')", true);
	    hkey.start();
    }
    
    function showDocument(id){    
		document.getElementById('ifrm' + id).style.width = "100%";
		document.getElementById('ifrm' + id).style.height = "600px";		
		document.getElementById('ifrm' + id).style.display = "block";	
		if (ifrmSel != id) {
		    hideDocument(ifrmSel);
		    ifrmSel = id;
            var hkey = new Hkey('hk' + id, 'interfacturas/interfacturasFirmaAjax.asp?id=' + id , '<% =HKEY() %>', 'firmar_callback()', true);
	        hkey.start();
	        loadInfoContable(id);
	        loadInfoImpositiva(id);
	        var tipoFactura = document.getElementById("tipoFac_" + id).value
	        if ( tipoFactura == '<%= FAC_CODIGO_CONCEPTO_GP %>') 
	            loadInfoAnalisis(id);
	        else
	            document.getElementById("infoAnalsisPuerto").innerHTML = "";
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
    function firmarAll_callback(cd, resp){  
        if (resp != "<%=RESPUESTA_OK%>"){	      
	        document.getElementById('hkAll' + cd).innerHTML =	resp;		
        } else {
            submitInfo();
        }	        
    }
    function firmar_callback(resp){
		if (resp != "<%=RESPUESTA_OK%>"){	
		    document.getElementById('hk' + ifrmSel).innerHTML =	resp + "<br>";			
			showDocument(ifrmSel);
		} else {
		    var idx = jQuery.inArray(ifrmSel + '', arr, 0); 
		    arr.splice(idx, 1);
		    if (arr.length > 0) {
		        showDocument(arr.slice(0, 1));
		        document.getElementById("tableDoc").deleteRow(idx+1);
		    } else {
		        document.getElementById("frmSel").submit();
		    }
		}
    }         
    	
    function loadInfoContable(id) {
            var ifrm = document.getElementById("ifrmContable");            
            ifrm.src = "interfacturas/interfacturasInfoContable.asp?id=" + id;            
    }
    function loadInfoImpositiva(id){
        ch.bind("interfacturas/interfacturasInfoImpositiva.asp?id=" +  id  ,"CallBack_verImpositiva("+id+")");
        ch.send();
    }
    
    function submitInfo(pg, lpp) {                
        document.getElementById("numeroPagina").value= pg;
        document.getElementById("registrosPorPagina").value= lpp;
        document.getElementById("frmSel").submit();
    }
    function loadInfoAnalisis(id) {
        ch.bind("interfacturas/interfacturasInfoAnalisis.asp?id=" +  id  ,"CallBack_verAnalisis("+id+")");
		ch.send();
    }
    function CallBack_verAnalisis(id){
		var respuesta = ch.response();
		document.getElementById("infoAnalsisPuerto").innerHTML = respuesta;
    }
    function CallBack_verImpositiva(id){
        var respuesta = ch.response();
        document.getElementById("infoImpositiva").innerHTML = respuesta;
    }
    function abrirAnalisis(pDtContable,pPto,pCartaPorte,pId,pTransporte,pCdProducto){
		if (pTransporte == <%=TIPO_TRANSPORTE_CAMION%>) var puw = new winPopUp('popupAnalisis','Puertos/InfoAnalisisCamion.asp?Pto='+pPto+'&dtContable='+pDtContable+'&ctaPorte='+pCartaPorte+'&camion='+pId,'780','550','Informacion de Camiones');
		if (pTransporte == <%=TIPO_TRANSPORTE_VAGON%>) var puw = new winPopUp('popupAnalisis','Puertos/Operativo/OperativosPopUp.asp?Pto='+pPto+'&fecha='+pDtContable+'&cartaPorte='+pCartaPorte+'&cdoperativo='+pId+'&cdProducto='+pCdProducto,'780','550','Informacion de Vagones');
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
	    <form method="post" name="frmSel" id="frmSel" action="interfacturasFirma.asp">	
	        <input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		    <input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>	
	    </form>	
		<div class="tableaside size50">			
			<table id="tableDoc" class="datagrid" width="90%" align="center">					
				<thead>					
					<tr>
					    <th class="thiconac" align="center" width="30%" nowrap>	<% =GF_TRADUCIR("Sector") %></th>
						<th class="thiconac" align="center" width="35%" nowrap>	<% =GF_TRADUCIR("Proforma") %></th>						
						<th class="thiconac" align="center" width="30%" nowrap>	<% =GF_TRADUCIR("Fecha") %></th>
						<th class="thiconac" align="center" width="20px" nowrap><% =GF_TRADUCIR("Firmar") %></th>
						<th class="thiconac" align="center" width="20px" nowrap><% =GF_TRADUCIR("Ver") %></th>
					</tr>
				</thead>	
				<tbody id="tbody"> 	
				<% if (not rs.eof) then
						while (not rs.eof) %>			
						<tr>
							<td align="center"><% =rs("NOSESC") %></td>
							<td align="center"><% =rs("FCCMNR") %></td>
							<td align="center"><% =GF_FN2DTE(rs("FCCMFC"))%></td>
							<td align="center"><div id="hk<%=rs("FCRGNR")%>"></div></td>
							<td align="center"><img src="images/search-16.png" style="cursor:pointer" onclick="javascript:showDocument('<%=rs("FCRGNR")%>')" /></td>
							<input type="hidden" id="tipoFac_<%=rs("FCRGNR")%>" name="tipoFac_<%=rs("FCRGNR")%>" value="<%=rs("LDLYCD")%>">
						</tr>
						<%	
							rs.MoveNext()
						wend
					else %>					
						<tr>
							<td colspan="4" align="center"><%=GF_TRADUCIR("No tiene facturas por firmar")%></td>
						</tr>
				<%	end if	%>
				</tbody>					
				
				<tfoot>
					<tr>
						<td colspan="4"><div id="paginacion"></div></td>						
					</tr>					
				</tfoot>				
			</table>
			<table id="table1" class="datagrid" width="50%" align="center">			
			    <tr>
			        <td>
			            Fimar todas las facturas de Acondicionamiento
			        </td>
			        <td width="20px" >
			            <br /><div id="hkAll<% =FAC_CODIGO_CONCEPTO_GP %>" style="text-align: center;"></div>
			        </td>
			    </tr>
			</table>
			<div> 
			    <iframe id='ifrmContable' width="100%" style="border: none;" src='' onload="autoAjustIframe(this)"></iframe>
			</div>
            <div id='infoImpositiva' width="100%"></div>
			<div id='infoAnalsisPuerto' width="100%"></div>
		</div>	
		
		<div id="ifrmDiv" class="tableaside size50"> </div>		
    </body>        
</html>