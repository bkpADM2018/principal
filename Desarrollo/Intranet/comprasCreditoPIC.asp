<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
dim idCotizacion, idArticulo, accion, ncrTotalPesos, ncrTotalDolares, saldoPesos, saldoDolares
dim guardado, aprobado, idArea, idDetalle

guardado = false
aprobado = false
idCotizacion = GF_Parametros7("idCotizacion",0,6)
idArticulo = GF_Parametros7("idArticulo",0,6)
idArea = GF_Parametros7("idArea",0,6)
idDetalle = GF_Parametros7("idDetalle",0,6)
accion = GF_PARAMETROS7("accion","",6)
ncrTotalPesos = GF_PARAMETROS7("ncrTotalPesos",2,6)		'El nuevo credito parcial ingresado en Pesos.
ncrTotalDolares = GF_PARAMETROS7("ncrTotalDolares",2,6) 'El nuevo credito parcial ingresado en Dolares.

if idCotizacion <> 0 then
	Call readCTZDetail(idCotizacion, idArticulo, idArea, idDetalle)	
end if

'El tipo de cambio se determina al momento de realizar tanto el PIC como cualquier pago, esto trae aparejado que 
'el mismo varía de pago en pago, generando así un desbalance entre los pesos y dolares de saldo de un PIC, de aquí
'que los controles al momento de cargar un pago se realicen solo sobre la moneda en la cual esta nominado el PIC.
'Al momento de autorizar un credito se debe evitar utilizar el tipo de cambio original dado que como ya se explicó, la 
'relación del saldo pesos-dolares no está regida por este valor. Para evitar situaciones incoherentes como la aparición de 
'importes negativos (en pesos o dolares), se determinará un tipo de cambio propio para el PIC segun los saldos restantes.
'Esto no tiene ningún inconveniete dado que en general el usuario que realiza un credito solo lo hace en una moneda, usualmente 
'la moneda de nominación del PIC.
if (ctz_det_importeDolaresFacturado > 0) then ctz_det_TipoCambio = round(ctz_det_importePesosFacturado/ctz_det_importeDolaresFacturado, 3)

if ((ncrTotalDolares <> 0) or (ncrTotalPesos <> 0)) then
	if ncrTotalDolares = 0 then
		if (ctz_det_TipoCambio <> 0) then ncrTotalDolares = round(ncrTotalPesos/ctz_det_TipoCambio,2)
	end if	
	if ncrTotalPesos = 0 then
		if (ctz_det_TipoCambio <> 0) then ncrTotalPesos = round(ncrTotalDolares*ctz_det_TipoCambio,2)
	end if	
end if

saldoPesos = ctz_det_importePesosFacturado-ctz_det_ImportePesosCredito
saldoDolares = ctz_det_importeDolaresFacturado-ctz_det_ImporteDolaresCredito

if Controlar() then
	if (accion = ACCION_GRABAR) then
		guardado = true
		aprobado = true	'No se necesita ninguna aprobacion del credito, pero se deja preparado por si en el futuro lo piden.
		'Se graba el importe habilitado de creditos.
		strSQL="Update TOEPFERDB.TBLCTZDETALLE set IMPORTEPESOSCREDITO= IMPORTEPESOSCREDITO + " & (ncrTotalPesos*100)
		strSQL= strSQL & ", IMPORTEDOLARESCREDITO= IMPORTEDOLARESCREDITO + " & (ncrTotalDolares*100)
		strSQL= strSQL & " where IDCOTIZACION=" & idCotizacion & " and IDARTICULO=" & idArticulo & " and IDAREA=" & idArea & " and IDDETALLE=" & idDetalle
		Call executeQuery(rs, "OPEN", strSQL)		
	end if		
end if

'----------------------------------------------------------------------------------------------------------------------------------
function Controlar()	
    if (ctz_cdMoneda = MONEDA_PESO) then
	    if (round((ncrTotalPesos*100), 0) > round(saldoPesos, 0)) then Call setError(CTZ_NCR_MAYOR_SALDO)	
	end if
	if (ctz_cdMoneda = MONEDA_DOLAR) then
	    if (round((ncrTotalDolares*100), 0) > round(saldoDolares, 0)) then	Call setError(CTZ_NCR_MAYOR_SALDO)	
	end if
	if not hayError() then Controlar = true
	
end function
'----------------------------------------------------------------------------------------------------------------------------------
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Credito PIC") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.labelStyle {
	font-weight: bold;	
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.msgOK {
	font-weight: bold;
	font-size: 14px;
	color: #44CC66;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript">
	var importesItemDolares;
	var importesItemPesos;
	var importesScrDolares;	//Backup de los valores enviados a pantalla (Sirve para detectar cambio)
	var importesScrPesos;		//Backup de los valores enviados a pantalla (Sirve para detectar cambio)
	var refpopupCreditoPIC;
	function bodyOnLoad(){
		var tb = new Toolbar('toolbar', 4, 'images/compras/');
		//tb.addButton("Home-16x16.png", "Home", "irHome()");		
		//tb.addButtonREFRESH("Recargar", "submitInfo()");
		<% if not guardado then %>
		idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
		idBtnControl = tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
		<% end if %>
		tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		tb.draw();
	}
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}		
	function cerrar(){
		refpopupCreditoPIC = startIWin('popupCreditoPIC');
		refpopupCreditoPIC.hide();	
	}	
function sumarTotal(cur) {
		var tipoCambio = <% =Replace(ctz_det_TipoCambio, ",", ".") %>;
		var totalDolares = 0;
		var totalPesos = 0;
		var aux1 = 0;

		var objP = document.getElementById("ncrTotalPesos");
		var objPValue = objP.value.replace(/,/,".");
		var objD = document.getElementById("ncrTotalDolares");
		var objDValue = objD.value.replace(/,/,".");
		if (cur == "D") {
			if (importesScrDolares != objD.value){ 
				aux1 = objD.value * tipoCambio;
				objP.value = aux1;
			}	
		} else {
			if (importesScrPesos != objP.value)
				if (tipoCambio > 0){ 
					aux1 = objPValue / tipoCambio;
					objD.value = aux1;					
				}	
		}				
		importesItemDolares = objD.value;
		importesItemPesos = objP.value;
		objP.value = editarImporte(objP.value);
		if (objP.value == 0) objP.value = "";
		objD.value = editarImporte(objD.value);
		if (objD.value == 0) objD.value = "";
		importesScrDolares = objD.value;
		importesScrPesos = objP.value;		
	}	

	
</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="comprasCreditoPIC.asp">
<div id="toolbar"></div><br>
<table class="reg_header" align="center" width="95%" border="0" >				
	<tr>
		<td colspan="3"><% call showErrors() %></td>
	</tr>
	<tr>
		<td align="right" class="numberStyle" colspan="3"><% =GF_TRADUCIR("Id PIC:") %>&nbsp;<% =ctz_IdCotizacion %></td>				
	</tr>
	<%call getArticuloFull(idArticulo, descArticulo, "")%>
	<tr>
		<td colspan="3"><b><%=idArticulo & " - " & trim(descArticulo) & "&nbsp;&nbsp;&nbsp;(" & idArea & "-" & idDetalle & ")"%></b></td>
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound" colspan="3"><% =GF_TRADUCIR("Datos del Pedido") %></td>				
	</tr>
	<tr>
		<td></td>				
		<td align="center" width="20%"><b><u><% =GF_TRADUCIR("Pesos")   %></u></b></td>
		<td align="center" width="20%"><b><u><% =GF_TRADUCIR("Dolares") %></u></b></td>
	</tr>
	<tr>
		<td COLSPAN="3" align="right"><HR></td>	
	</tr>			
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Facturado hasta el momento") %></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ctz_det_importePesosFacturado,2)%></b></font></td>
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ctz_det_importeDolaresFacturado,2)%></b></font></td>	
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Créditos Pendientes") %></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ctz_det_importePesosCredito,2)%></b></font></td>
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ctz_det_importeDolaresCredito,2)%></b></font></td>
	</tr>
	<tr>
		<td class="reg_headeAr_nav recuadroRound"></td>				
		<td align="right"><HR></td>	
		<td align="right"><HR></td>	
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><b><% =GF_TRADUCIR("Saldo") %></b></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(saldoPesos,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(saldoDolares,2)%></b></font></td>	
	</tr>
	<tr>
		<td COLSPAN="3" align="right"><HR></td>	
	</tr>
	<%if (not guardado) then %>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Adicionar Nuevo Crédito del Articulo") %></td>				
		<td align="right"><input style="text-align:right;" type="text" onBlur="sumarTotal('P')" name="ncrTotalPesos" id="ncrTotalPesos" size="10" onkeypress="return controlIngreso(this, event, 'I');" value="<%=ncrTotalPesos%>"></td>
		<td align="right"><input style="text-align:right;" type="text" onBlur="sumarTotal('D')" name="ncrTotalDolares" id="ncrTotalDolares" size="10" onkeypress="return controlIngreso(this, event, 'I');" value="<%=ncrTotalDolares%>"></td>
	</tr>
	<tr>
		<td COLSPAN="3" align="right"><HR></td>	
	</tr>
	<%end if %>
</table>
<input type="hidden" name="accion" id="accion" value="<% =ACCION_CONTROLAR %>">
<input type="hidden" name="idCotizacion" id="idCotizacion" value="<%=idCotizacion %>">
<input type="hidden" name="idArticulo" id="idArticulo" value="<%=idArticulo %>">
<input type="hidden" name="idArea" id="idArea" value="<% =idArea %>">
<input type="hidden" name="idDetalle" id="idDetalle" value="<% =idDetalle %>">
</form>
</body>
</html>