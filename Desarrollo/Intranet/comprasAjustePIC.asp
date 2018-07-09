<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%

'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************

dim idCotizacion, idArticulo, accion, newTotalPesos, newTotalDolares, saldoPesos, saldoDolares, ajuPesos, ajuDolares, ctzAjuComentario
dim guardado, idArea, idDetalle, abrevUnidad, ajuCantidad, saldoCantidad, estado, importeaju, tipoInputPesos, tipoInputDolares, cdSolicitante
dim ajuImporte

guardado = false
idCotizacion = GF_Parametros7("idCotizacion",0,6)
idArticulo = GF_Parametros7("idArticulo",0,6)
idArea = GF_Parametros7("idArea",0,6)
idDetalle = GF_Parametros7("idDetalle",0,6)
accion = GF_PARAMETROS7("accion","",6)
newTotalPesos = GF_PARAMETROS7("newTotalPesos",2,6)
newTotalDolares = GF_PARAMETROS7("newTotalDolares",2,6)
newTotalCantidad = GF_PARAMETROS7("newTotalCantidad",2,6)
ctzAjuComentario = GF_PARAMETROS7("ctzAjuComentario","",6)

if idCotizacion <> 0 then
	Call readCTZDetail(idCotizacion, idArticulo, idArea, idDetalle)
	'Si no se encontró el item y el mismo no puede cargarse directamente en el PIC, permitir su carga por ajuste. (Esto solo se aplica para items muy especiales. 	
	if ((ctz_IdCotizacion = 0) and (not esArticuloElegiblePIC(idArticulo))) then
	    Call setUpSpecialItems(idCotizacion, idArticulo)		
	    flagIsNew = true
	end if
end if

ctz_det_TipoCambio = getTipoCambio(MONEDA_DOLAR,"")
ctz_det_importePesosFacturado = ctz_det_importePesosFacturado + ctz_det_ImportePesosCredito
ctz_det_importeDolaresFacturado = ctz_det_importeDolaresFacturado + ctz_det_ImporteDolaresCredito
'El tipo de cambio se determina al momento de realizar tanto el PIC como cualquier pago, esto trae aparejado que 
'el mismo varía de pago en pago, generando así un desbalance entre los pesos y dolares de saldo de un PIC, de aquí
'que los controles al momento de cargar un pago se realicen solo sobre la moneda en la cual esta nominado el PIC.
'Al momento de realizar un ajuste se debe evitar utilizar el tipo de cambio original dado que como ya se explicó, la 
'relación del saldo pesos-dolares no está regida por este valor. Para evitar situaciones incoherentes como la aparición de 
'importes negativos (en pesos o dolares), se determinará un tipo de cambio propio para el PIC segun los saldos restantes.
'Esto no tiene ningún inconveniete dado que en general el usuario que realiza un ajuste solo lo hace en una moneda, usualmente 
'la moneda de nominación del PIC.
if ((ctz_det_importeDolaresFacturado > 0) and (ctz_det_importePesosFacturado > 0))then ctz_det_TipoCambio = round(ctz_det_importePesosFacturado/ctz_det_importeDolaresFacturado, 3)


if ((newTotalDolares = 0) and (newTotalPesos = 0) and (newTotalCantidad = 0) and (accion = "")) then
	newTotalPesos = ctz_det_importePesos/100
	newTotalDolares = ctz_det_importeDolares/100
	newTotalCantidad = ctz_det_ArticuloCantidad
else	
	if newTotalDolares = 0 then
		if (ctz_det_TipoCambio <> 0) then newTotalDolares = round(newTotalPesos/ctz_det_TipoCambio,2)
	end if	
	if newTotalPesos = 0 then
		if (ctz_det_TipoCambio <> 0) then newTotalPesos = round(newTotalDolares*ctz_det_TipoCambio,2)
	end if	
end if

'Si se anula completamente uno de los dos importes, se anula el otro.
if (((newTotalPesos*100) = ctz_det_importePesosFacturado) or ((newTotalDolares*100) = ctz_det_importeDolaresFacturado)) then
	newTotalPesos = ctz_det_importePesosFacturado/100
	newTotalDolares = ctz_det_importeDolaresFacturado/100
end if

ajuPesos = round((newTotalPesos*100) - ctz_det_importePesos, 0)
ajuDolares = round((newTotalDolares*100) - ctz_det_importeDolares, 0)
ajuCantidad = newTotalCantidad - ctz_det_ArticuloCantidad
ajuImporte = ajuPesos
if (CTZ_cdMoneda = MONEDA_DOLAR) then ajuImporte = ajuDolares
saldoPesos = ctz_det_importePesos-ctz_det_importePesosFacturado+ajuPesos
saldoDolares = ctz_det_importeDolares-ctz_det_importeDolaresFacturado+ajuDolares
saldoCantidad = ctz_det_ArticuloCantidad-ctz_det_Facturado+ajuCantidad

if Controlar then
	if accion = ACCION_GRABAR then
	
	    if (flagIsNew) then Call addCTZItems(idCotizacion, idArticulo, newTotalCantidad, ctz_det_ArticuloIdUnidad, idArea, idDetalle, ctz_det_importePesos, ctz_det_importeDolares, ctz_det_TipoCambio)
	    
		guardado = true
		'GUARDAR <CABECERA> DEL AJUSTE
		strSQL = "SELECT * FROM TBLCTZAJUSTES WHERE IDCOTIZACION=" & idCotizacion & " AND IDARTICULO=" & idArticulo & " AND IDAREA=" & idArea & " AND IDDETALLE=" & idDetalle & " AND APLICADO='" & TIPO_NEGACION & "'" 
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then
			myIdAjuste = rs("IDAJUSTE")
			strSQL = "UPDATE TBLCTZAJUSTES SET APLICADO='" & TIPO_NEGACION & "', IMPORTEPESOS=" & ajuPesos & ", IMPORTEDOLARES=" & ajuDolares & ", CANTIDAD=" & ajuCantidad & ", OBSERVACIONES='" & ctzAjuComentario & "', CDUSUARIO='" & session("usuario") & "', MOMENTO=" & session("MmtoSistema") & " WHERE IDAJUSTE=" & myIdAjuste
	        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		else
			strSQL = "INSERT INTO TBLCTZAJUSTES (IDCOTIZACION, IDARTICULO, IDAREA, IDDETALLE, IMPORTEPESOS, IMPORTEDOLARES, CANTIDAD, OBSERVACIONES, APLICADO, CDUSUARIO, MOMENTO) VALUES(" & idCotizacion & "," & idArticulo & "," & idArea & "," & idDetalle & "," & ajuPesos & "," & ajuDolares & ", " & ajuCantidad & ", '" & ctzAjuComentario & "','" & TIPO_NEGACION & "','" & session("usuario") & "'," & session("MmtoSistema") & ")"		
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)
			strSQL = "SELECT * FROM TBLCTZAJUSTES WHERE IDCOTIZACION=" & idCotizacion & " AND IDARTICULO=" & idArticulo & " AND IDARTICULO=" & idArticulo & " AND IDAREA=" & idArea & " AND IDDETALLE=" & idDetalle & " AND APLICADO='" & TIPO_NEGACION & "'" 
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			myIdAjuste = rs("IDAJUSTE")
		end if	
		
		'Response.Write strSQL

		'GUARDAR <FIRMAS> DEL AJUSTE
		'Si el precio unitario de la mercadería aumentó, se debe autorizar el ajuste.
		if (ctz_cdMoneda = MONEDA_PESO) then
			precioUOriginal = Round(ctz_det_importePesos/ctz_det_ArticuloCantidad, 2)
			precioUNuevo = 0 
			if (saldoCantidad > 0) then precioUNuevo = Round(saldoPesos/saldoCantidad, 2)
			importeaju = ajuPesos
		else
			precioUOriginal = Round(ctz_det_importeDolares/ctz_det_ArticuloCantidad, 2)
			precioUNuevo = 0 
			if (saldoCantidad > 0) then precioUNuevo = Round(saldoDolares/saldoCantidad, 2)
			importeaju = ajuDolares
		end if
		'Si el precio promedio original no se ve incrementado, se da por autorizado el ajuste dado que se restringe el pago a realizar.		
		if ((precioUOriginal >= precioUNuevo) and (importeaju <= 0)) then

			'ACTUALIZAR DETALLE PIC, CARGAR NUEVO DETALLE PARA QUE COINCIDAN LOS IMPORTES
		    strSqlNuevoItem = "UPDATE TBLCTZDETALLE SET IMPORTEPESOS=" & cdbl(newTotalPesos*100) & " , IMPORTEDOLARES=" & cdbl(newTotalDolares*100) & ", CANTIDAD=" & newTotalCantidad & " WHERE IDCOTIZACION=" & idCotizacion & " AND IDARTICULO=" & idArticulo & " AND IDAREA=" & idArea & " AND IDDETALLE=" & idDetalle
		    'Response.Write strSqlNuevoItem
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSqlNuevoItem)
			
			'Se recalcula la cantidad facturada dado que al cambiar el importe o la cantidad original, lo ya facturado cambia.
			if (ctz_cdMoneda = MONEDA_PESO) then
				aux = "(CANTIDAD*IMPORTEPESOSFACTURADO)/IMPORTEPESOS"
				if (newTotalPesos = 0) then aux = "0"
				aux2 = "IMPORTEPESOSFACTURADO <> IMPORTEPESOS"								
			else
				aux = "(CANTIDAD*IMPORTEDOLARESFACTURADO)/IMPORTEDOLARES"
				if (newTotalDolares = 0) then aux = "0"
				aux2 = "IMPORTEDOLARESFACTURADO <> IMPORTEDOLARES"				
			end if										
			strSQLAux="Update TBLCTZDETALLE set FACTURADO=" & aux & " where IDCOTIZACION=" & idCotizacion & " AND IDARTICULO=" & idArticulo & " AND IDAREA=" & idArea & " AND IDDETALLE=" & idDetalle
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQLAux)
			'SE DETERMINA EL NUEVO ESTADO DEL PIC.
		    strSQLAux="Select count(*) CANT from TBLCTZDETALLE where " & aux2 & " and IDCOTIZACION=" & idCotizacion			
		    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQLAux)
			estado= CTZ_FIRMADA
			if (rs("CANT") = 0) then estado= CTZ_FACTURADA
			
			
			'ACTUALIZAR CABECERA PIC, MODIFICAR TOTALES
			strSQLAux = "UPDATE TBLCTZCABECERA SET ESTADO='" & estado & "', IMPORTEPESOS=" & cdbl(ctz_ImportePesos) + cdbl(ajuPesos) & ", IMPORTEDOLARES=" & cdbl(ctz_ImporteDolares) + cdbl(ajuDolares) & " WHERE IDCOTIZACION=" & idCotizacion
			'Response.Write strSQLAux
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQLAux)

			'ACTUALIZAR ESTADO DE AJUSTE			
			strSQLAux = "UPDATE TBLCTZAJUSTES SET APLICADO='" & TIPO_AFIRMACION & "' WHERE IDAJUSTE=" & myIdAjuste
			'Response.Write strSQLAux
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQLAux)
			Call setInfo(CTZ_AJU_APROBADO)		
		else 			
			'CAMBIAR ESTADO PIC, EL PIC PASA A ESTADO EN AJUSTE
			strSQLAux = "UPDATE TBLCTZCABECERA SET ESTADO='" & CTZ_EN_AJUSTE & "' WHERE IDCOTIZACION=" & idCotizacion			
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQLAux)
			'Se cargan las firmas
			strSqlFirmas = "SELECT CDUSUARIO FROM TBLCTZFIRMAS WHERE IDCOTIZACION = " & idCotizacion & " and SECUENCIA in (" & PIC_FIRMA_RESPONSABLE & ", " & PIC_FIRMA_GTE_SECTOR & ")"
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSqlFirmas)
			if (not rs.eof) then
			    cdSolicitante=rs("CDUSUARIO")			    			    
			    rs.MoveNext()
			    if (not rs.eof) then cdAutorizante=rs("CDUSUARIO")			    
            end if			    
            
			Call addAJUCTZFirmas(idCotizacion, myIdAjuste, cdSolicitante, cdAutorizante)
            
			Call setWarning(CTZ_AJU_NO_APROBADO)			
		end if							
	end if		
end if
'----------------------------------------------------------------------------------------------------------------------------------
'Carga de datos iniciales para los items especiales que no existen la primera vez que sae intenta ajustarlos.
'Estos items son sumamente reros y por esoi tienen un permiso especial para ajustarse ya que no se los puede elegir al cargar un PIC.
Function setUpSpecialItems(idCotizacion, idArticulo)

    Dim cdUnidad, dsUnidad
    
    Select case (idArticulo)
        Case CTZ_ITEM_DIFF_CAMBIO:
            ctz_IdCotizacion = idCotizacion
	    	ctz_det_IdArticulo = idArticulo
	    	ctz_det_TipoCambio = getTipoCambio(MONEDA_DOLAR, "")
	    	ctz_cdMoneda = MONEDA_DOLAR
	    	ctz_det_ArticuloCantidad = 1	    	
	    	Call getUnidadArticulo(idArticulo, ctz_det_ArticuloIdUnidad, cdUnidad, dsUnidad)
	    	ctzAjuComentario = "Diferencia de Cambio"	    	
	End Select
	
End Function
'----------------------------------------------------------------------------------------------------------------------------------
function Controlar
if accion = "" then 
	accion = ACCION_CONTROLAR
else	
    if ((CTZ_idContrato > 0) and (ajuImporte > 0)) then
        Call setError(CTZ_AJS_POS_NO_PERMITIDO)
    else
        if (ctz_cdMoneda = MONEDA_PESO) then
	        if (Round(cdbl(newTotalPesos)*100, 0) < Round(cdbl(ctz_det_importePesosFacturado), 0)) then Call setError(CTZ_AJU_TOTAL_BAJO)
	    else
	        if (Round(cdbl(newTotalDolares)*100, 0) < Round(cdbl(ctz_det_importeDolaresFacturado), 0)) then Call setError(CTZ_AJU_TOTAL_BAJO)
	    end if
    	
	    if (newTotalCantidad < ctz_det_Facturado) then Call setError(CTZ_AJU_CANT_TOTAL_BAJO)
    	
	    if  ((cdbl(newTotalPesos)*100 = cdbl(ctz_det_importePesos)) and (cdbl(newTotalDolares)*100 = cdbl(ctz_det_importeDolares)) and (newTotalCantidad= ctz_det_ArticuloCantidad)) then
		    Call setError(CTZ_AJU_IMP_IGUALES)
	    end if			
	    if len(ctzAjuComentario) < 1 then Call setError(COMENTARIO_REQUERIDO) 
	end if
end if
if not hayError() then Controlar = true
end function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Ajuste " & ctz_docCode) %></title>
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
	var refpopupAjuPIC;
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
		refpopupAjuPIC = startIWin('popupAjuPIC');
		refpopupAjuPIC.hide();	
	}	
function sumarTotal(cur) {
		var tipoCambio = <% =Replace(ctz_det_TipoCambio, ",", ".") %>;
		var totalDolares = 0;
		var totalPesos = 0;
		var aux1 = 0;

		var objP = document.getElementById("newTotalPesos");
		var objPValue = objP.value.replace(/,/,".");
		var objD = document.getElementById("newTotalDolares");
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
<form method="post" id="frmSel" action="comprasAjustePIC.asp?idCotizacion=<%=idCotizacion%>&idArticulo=<%=idArticulo%>">
<div id="toolbar"></div><br>
<table class="reg_header" align="center" width="95%" border="0">				
	<tr>
		<td colspan="5"><% call showErrors() %></td>
	</tr>
	<tr>
		<td align="right" class="numberStyle" colspan="5"><% =GF_TRADUCIR("Id " & ctz_docCode & ":") %>&nbsp;<% =ctz_IdCotizacion %></td>				
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound" colspan="5"><% =GF_TRADUCIR("Datos del Pedido") %></td>
	</tr>
	<tr>
		<td></td>						
		<td align="center" width="20%" id="tituloPesos"><b><u><% =GF_TRADUCIR("Pesos")   %></u></b></td>
		<td align="center" width="20%" id="tituloDolares"><b><u><% =GF_TRADUCIR("Dolares") %></u></b></td>				
		<td align="center" width="20%" id="tituloPesos" colspan="2"><b><u><% =GF_TRADUCIR("Cantidad")   %></u></b></td>		
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Total del pedido") %></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ctz_importePesos,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ctz_importeDolares,2)%></b></font></td>	
	</tr>
	<tr>
		<td COLSPAN="5" align="right"><HR></td>	
	</tr>
	<% Call getArticuloFull(idArticulo, descArticulo, abrevUnidad)%>
	<tr>
		<td colspan="3"><b><%=idArticulo & " - " & trim(descArticulo) & "&nbsp;&nbsp;&nbsp;(" & idArea & "-" & idDetalle & ")"%></b></td>	
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Total del articulo") %></td>				
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ctz_det_importePesos,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ctz_det_importeDolares,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=GF_EDIT_DECIMALS(ctz_det_ArticuloCantidad*100, 2)%></b></font></td>		
		<td><% = abrevUnidad %></td>
	</tr>	
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Facturado/Recibido hasta el momento") %></td>						
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ctz_det_importePesosFacturado,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ctz_det_importeDolaresFacturado,2)%></b></font></td>			
		<td align="right"><font size="+1"><b><%=GF_EDIT_DECIMALS(ctz_det_Facturado*100, 2)%></b></font></td>		
		<td><% = abrevUnidad %></td>
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Ajuste") %></td>						
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(ajuPesos,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(ajuDolares,2)%></b></font></td>			
		<td align="right"><font size="+1"><b><%= GF_EDIT_DECIMALS(ajuCantidad*100, 2)%></b></font></td>
		<td><% = abrevUnidad %></td>
	</tr>
	<tr>
		<td class="reg_headeAr_nav recuadroRound"></td>				
		<td align="right"><HR></td>	
		<td align="right"><HR></td>
		<td align="right" colspan="2"><HR></td>		
	</tr>
	<tr>
		<td class="reg_header_nav recuadroRound"><b><% =GF_TRADUCIR("Saldo") %></b></td>		
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(saldoPesos,2)%></b></font></td>	
		<td align="right"><font size="+1"><b><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(saldoDolares,2)%></b></font></td>			
		<td align="right"><font size="+1"><b><%=GF_EDIT_DECIMALS(saldoCantidad*100, 2)%></b></font></td>
		<td><% = abrevUnidad %></td>
	</tr>
	<tr>
		<td COLSPAN="5" align="right"><HR></td>	
	</tr>
	<%if not guardado then %>
	<tr>
		<td class="reg_header_nav recuadroRound"><% =GF_TRADUCIR("Nuevo Total del Articulo") %></td>				
		<%	if (ctz_cdMoneda = MONEDA_PESO) then 
				tipoInputPesos = "text" 
				tipoInputDolares = "hidden"
			else
				tipoInputPesos = "hidden" 
				tipoInputDolares = "text"
			end if
		%>
		<td align="right"><input style="text-align:right;" type="<% =tipoInputPesos %>" onBlur="sumarTotal('P')" name="newTotalPesos" id="newTotalPesos" size="10" onkeypress="return controlIngreso(this, event, 'I');" value="<%=newTotalPesos%>"></td>
		<td align="right"><input style="text-align:right;" type="<% =tipoInputDolares %>" onBlur="sumarTotal('D')" name="newTotalDolares" id="newTotalDolares" size="10" onkeypress="return controlIngreso(this, event, 'I');" value="<%=newTotalDolares%>"></td>		
		<td align="right"><input style="text-align:right;" type="text" name="newTotalCantidad" size="10" value="<% =newTotalCantidad %>"></td>
		<td><% = abrevUnidad %></td>
	</tr>
	<tr>
		<td COLSPAN="5" align="right"><HR></td>	
	</tr>
	<%end if %>
	<tr>
		<td class="reg_header_nav recuadroRound" colspan="5"><% =GF_TRADUCIR("Justificación del Ajuste") %></td>				
	</tr>
	<tr>		
			<%if guardado then			
				Response.Write "<td colspan='3' align=left>" & ctzAjuComentario
			else%>	
				<td colspan="5" align=center>
					<textarea name="ctzAjuComentario" id="ctzAjuComentario" cols="100"><%=ctzAjuComentario%></textarea>				
			<%end if%>	
		</td>
	</tr>
</table>
<input type="hidden" name="accion" id="accion">
<input type="hidden" name="idArticulo" id="idArticulo">
<input type="hidden" name="idArea" id="idArea" value="<% =idArea %>">
<input type="hidden" name="idDetalle" id="idDetalle" value="<% =idDetalle %>">
</form>
</body>
</html>