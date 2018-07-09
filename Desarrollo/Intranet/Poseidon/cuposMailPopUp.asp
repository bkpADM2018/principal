<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<%

'********************************************************************************************************************************
'********************************************************* INICIO DE PAGINA *****************************************************
'********************************************************************************************************************************
Dim cuitCupeador, cdProducto, nroPuerto, fechaDesde, fechaHasta, idCorredor , idVendedor, rsNom, g_strPuerto,fechaPermitida, totalNominados, cuitCliente
Dim strSQL, rs, receptor, receptorDS, action, msg, auxDestino
Dim hayCorredor, hayVendedor, fc

cuitCupeador = GF_PARAMETROS7("cuitCupeador",0,6)
cuitCliente = GF_PARAMETROS7("cuitCliente",0,6)
cdProducto  = GF_PARAMETROS7("cdProducto",0,6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
fechaDesde  = GF_PARAMETROS7("fechaDesde","",6)
fechaHasta  = GF_PARAMETROS7("fechaHasta","",6)
idCorredor  = GF_PARAMETROS7("cdCorredor","",6)
idVendedor  = GF_PARAMETROS7("cdVendedor","",6)
action      = GF_PARAMETROS7("action","",6)
mails       = GF_PARAMETROS7("mails","",6)
mrecep		= GF_PARAMETROS7("mrecep","",6)
fc	        = GF_PARAMETROS7("fc","",6)

'Obtengo lel destinatario de los mails.
hayCorredor = False
if ((idCorredor <> "") and (CLng(idCorredor) <> 0) and  (Clng(idCorredor) <> SIN_CORREDOR)) then hayCorredor = True
hayVendedor = False
if ((idVendedor <> "") and (CLng(idVendedor) <> 0) and  (Clng(idVendedor) <> SIN_CORREDOR)) then hayVendedor = True

'--- Determino a quien se envia el mail ---
receptor = cuitCliente
receptorDs = getDsClienteByCUIT(cuitCliente)
if (mrecep = "") then	
	'No se indica un receptor, se sugiere uno en funcion de la situacion de la operatoria.
	if (CDbl(cuitCupeador) = CDbl(cuitCliente)) then 		
		'Si el que cupea es el cliente/destinatario de la mercaderia
		if (hayCorredor) then			
			receptor = getCuitCorredorByCd(g_strPuerto, idCorredor)
			receptorDs = getDsCorredor(idCorredor)
			mrecep = "R"
		else
			if (hayVendedor) then
				receptor = getCuitVendedorByCd(g_strPuerto, idVendedor)
				receptorDs = getDsVendedor(idVendedor)
				mrecep = "V"
			end if
		end if    		
	else
		if (CDbl(cuitCupeador) <> CDbl(CUIT_TOEPFER)) then
			'Si el que cupea es el corredor (y no es TOEPFER ni el cliente/destinatario)
			if (hayVendedor) then
				receptor = getCuitVendedorByCd(g_strPuerto, idVendedor)
				receptorDs = getDsVendedor(idVendedor)
				mrecep = "V"
			end if
		end if
	end if      
else
	if (hayCorredor and mrecep = "R") then
        receptor = getCuitCorredorByCd(g_strPuerto, idCorredor)
        receptorDs = getDsCorredor(idCorredor)
		mrecep = "R"
	end if    
	if (hayVendedor and mrecep = "V") then
		receptor = getCuitVendedorByCd(g_strPuerto, idVendedor)
		receptorDs = getDsVendedor(idVendedor)
		mrecep = "V"
	end if
end if
'-----------------------------------------
if (action = ACCION_GRABAR) then 
    msg = "Las direcciones fueron guardadas correctamente"	
    Call saveTaskMailList(TASK_POS_ADMIN_CUPOS, receptor, mails)
	auxDestino = mails
else
	'Obtengo los mails del receptor.    
	auxDestino = getTaskMailList(TASK_POS_ADMIN_CUPOS, receptor)
end if

%>
<html>
<head>
<title>Sistema de Cupos - Envio de Mails</title>

<meta http-equiv="x-ua-compatible" content="IE=11">

<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
    
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript">
    var ch = new channel();

    function bodyOnLoad() {
    <%  if (msg <> "") then %>
        showMsg('<% =msg %>');
    <%  end if %>	
    }       
    
	function submitFrm(pAcc) {
		var frm = document.getElementById("frmSel");
		document.getElementById("action").value = pAcc;        
		frm.submit();
	}
	
    function enviarMail() {        
        var chk = document.getElementById("forzar");
        var forzar = 0;
        if (chk.checked) forzar = 1;
        var strParameter = "accion=<%=ACCION_EMAIL %>&cuitCupeador=<% =cuitCupeador %>&cuitCliente=<% =cuitCliente %>&fechaDesde=<% =fechaDesde %>&fechaHasta=<% =fechaHasta %>&cdVendedor=<% =idVendedor %>&cdCorredor=<% =idCorredor %>&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&mrecep=<% =mrecep %>&forzar=" + forzar + "&fc=<% =fc %>";        
        ch.bind("cuposAdministrarAjax.asp?" + strParameter, "enviarMail_callback()");
        ch.send();
    }
    
    function enviarMail_callback() {
        var resp = ch.response();		
        showMsg(resp);
    }
    
    function showMsg(pMsg) {
        document.getElementById("dsError").className = "TDSUCCESS";        
        document.getElementById("dsError").innerHTML = pMsg;
        $("#dsError").removeAttr("style");
        var om = document.getElementById("dsError");
        setTimeout(function() { om.style.display = "none"; } , 5000)
    }
    
</script>
</head>
<body onload="bodyOnLoad()">
    <form name="frmSel" id="frmSel" method="post" action="cuposMailPopUp.asp">
    <table cellpadding="2" cellspacing="1" >
        <tr>
            <td colspan="2">
                <div class="tableasidecontent">
    
                    <div class="col26 reg_header_navdos"> Fecha </div>
                    <div class="col46">Desde: <% =GF_FN2DTE(fechaDesde) %> Hasta: <% =GF_FN2DTE(fechaHasta) %> </div>
                            
                    <div class="col26 reg_header_navdos"> Producto </div>
                    <div class="col46"><%= cdProducto &"-"& Trim(getDsProducto(cdProducto)) %></div>
                    
                    <div class="col26 reg_header_navdos"> Destinatario </div>
                    <div class="col46" title="<% =GF_STR2CUIT(cuitCliente) %>"><% =getDsClienteByCUIT(cuitCliente) %></div>
                                          
                    <div class="col26 reg_header_navdos"> Corredor </div>
                    <div class="col46" title="<% =idCorredor %>"><% if (idCorredor > 0) then
                                            response.Write Trim(getDsCorredor(idCorredor)) 
                                          end if %></div>
                    
                    <div class="col26 reg_header_navdos"> Vendedor </div>
                    <div class="col46" title="<% =idVendedor %>"><% if (idVendedor > 0) then
                                            response.Write getDsVendedor(idVendedor) 
                                          end if %></div>                   
                </div>    
            
            </td>            
        </tr>
		<tr><td colspan="2"><hr></td></tr>
		<tr>
			<td class="reg_header_navdos">  
				Enviar mail a
            </td>
			<td>
				<input type="radio" name="mrecep" id="mrecepC" value="C" <% if (mrecep="C" or mrecep="") then response.write "checked" %> onclick="submitFrm('')"> Destinatario
				<% if (hayCorredor) then %> <input type="radio" name="mrecep" id="mrecepR" value="R" <% if (mrecep="R") then response.write "checked" %> onclick="submitFrm('')"> Corredor <% end if %>
				<% if (hayVendedor) then %> <input type="radio" name="mrecep" id="mrecepV" value="V" <% if (mrecep="V") then response.write "checked" %> onclick="submitFrm('')"> Vendedor <% end if %>
			</td>
		</tr>
		<tr><td colspan="2"><hr></td></tr>
		<tr>			
			<td style="font-weight:bold; font-size: 14px; text-align:center" colspan="2" title="<% =GF_STR2CUIT(receptor) %>"> <% =receptorDs %> </td>
		</tr>		
		<tr>
            <td colspan="16" id="tdError" >
                <div id="dsError" style="display:none;"></div>
            </td>
        </tr>   
		<tr>
			<td class="reg_header_navdos">  
				 <input type="checkbox" id="forzar" value="1" />
             </td>
			<td> Enviar Cupos ya informados. </td>
		</tr>
        <tr>
            <td class="col26 reg_header_navdos"> Mails: </td>
            <td> 
                <textarea name="mails" cols="50" rows="5"><% =Trim(auxDestino) %></textarea>                
                Separar las direcciones con punto y coma (;)
            </td>
        </tr>
		<tr><td>&nbsp;</td></tr>
        <tr>
            <td  colspan="2" align="center">
                <input type="button" value="Grabar" onclick="javascript:submitFrm('<% =ACCION_GRABAR %>')"/>&nbsp;&nbsp;
                <input type="button" value="Enviar" onclick="javascript:enviarMail()" />
            </td>            
        </tr>            
    </table>
    <input type="hidden" name="cuitCupeador" value="<% =cuitCupeador %>" />
    <input type="hidden" name="cuitCliente" value="<% =cuitCliente %>" />
    <input type="hidden" name="cdProducto" value="<% =cdProducto %>" />
    <input type="hidden" name="pto" value="<% =g_strPuerto %>" />
    <input type="hidden" name="fechaDesde" value="<% =fechaDesde %>" />
    <input type="hidden" name="fechaHasta" value="<% =fechaHasta %>" />
    <input type="hidden" name="cdCorredor" value="<% =idCorredor %>" />
    <input type="hidden" name="cdVendedor" value="<% =idVendedor %>" />
    <input type="hidden" name="action" id="action" value="<% =ACCION_GRABAR %>" />
    <input type="hidden" name="fc" id="fc" value="<% =fc %>" />
    </form>
       
</body>
</html>