<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Const TODOS_ARTICULOS = 0
dim CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraCuentaDS, CAB_ObraImporte, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada
dim idPedido, CAB_idCotizacion, CAB_idPedido, CAB_cdPedido, CAB_idProveedor, CAB_dsProveedor, CAB_fecEntrega, CAB_observaciones, CAB_importePesos, CAB_importeDolares
dim cantArt, index, accion, proveedoresL, idCotizacion, CAB_titulo, CAB_idObra, CAB_firmaSolicitante, CAB_firmaResponsable, CAB_firmaSupervisor,CAB_idContrato
dim IT_artID, IT_artDS, IT_cantidad, IT_importePesos, IT_importeDolares, IT_unidadID, IT_unidadDS, idCotizacionElegida, CAB_Importe
dim CAB_Moneda, CAB_IdDivision, CAB_ImportePlanilla, CAB_FechaBudget, IT_artBA, IT_artBD, IT_unidadCD, errFirma, IT_Importe
dim CAB_dsSolicitante, CAB_dsResponsable, CAB_dsSupervisor, CAB_idSupervisor, CAB_cdSolicitante, CAB_cdResponsable, CAB_cdSupervisor
dim firmante1Ds, firmante1Tx, firmante2Ds, firmante2Tx, firmante3Ds, firmante3Tx, firmante1Cd, firmante2Cd, firmante3Cd, firmante5Ds, firmante5Tx, firmante5Cd, firmante4Ds, firmante4Tx, firmante4Cd
Dim firmante1Rol, firmante2Rol, firmante3Rol, firmante4Rol, firmante5Rol
Dim firmante1Sec, firmante2Sec, firmante3Sec, firmante4Sec, firmante5Sec
dim dicArtUCC,bloqueoPorAFE,strObservaciones,flagDetallePresupuesto, totalImporteCompraUSD, mmtoDesde, CAB_Momento, cantPICDirectos365, montoPICDirectos365
Dim limiteCD, limiteSP, importeCompra, unidadCD, unidadSP, tituloFirma,CAB_estado,rolUsuario, cantPICDirectos30, montoPICDirectos30

Set dicArtUCC = Server.CreateObject("Scripting.Dictionary")
bloqueoPorAFE = false
flagDetallePresupuesto = false
'-----------------------------------------------------------------------------------------------
sub RedimVarialbles(pCant)
	redim IT_cantidad(pCant)
	redim IT_importePesos(pCant)
	redim IT_importeDolares(pCant)
	redim IT_unidadID(pCant)
	redim IT_unidadCD(pCant)
	redim IT_unidadDS(pCant)	
	redim IT_artDS(pCant)
	redim IT_artID(pCant)
	redim IT_artBA(pCant)
	redim IT_artBD(pCant)
end sub
'-------------------------------------------------------------------
Function cargarDetalles(pIdCotizacion)
	Dim strSQL, rsDET, connDET
		
	strSQL="SELECT * from TBLCTZDETALLE where IDCOTIZACION=" & idCotizacionElegida
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rsDET, "OPEN", strSQL)
	'Se inicializan las variables de artículos que estarán vacias inicialmente
	if (not rsDET.eof) then RedimVarialbles(rsDET.RecordCount)	
	while (not rsDET.eof)
		IT_artID(index) = rsDET("IDARTICULO")
		call getArticuloFull(IT_artID(index), IT_artDS(index), IT_unidadDS(index))
		IT_cantidad(index) = CDbl(rsDET("CANTIDAD"))
		IT_importePesos(index) = rsDET("IMPORTEPESOS")				
		IT_importeDolares(index) = rsDET("IMPORTEDOLARES")				
		IT_unidadID(index) = rsDET("IDUNIDAD")
		IT_artBA(index) = rsDET("IDAREA")
		IT_artBD(index) = rsDET("IDDETALLE")
		index = index + 1
		rsDET.movenext
	wend			
End Function
'-------------------------------------------------------------------
Function cargarFirmas(pIdCotizacion)
	Dim rsFirmas, connFirmas
	
	Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLCTZFIRMAS_GET_BY_IDCOTIZACION", pIdCotizacion)
	while not rsFirmas.eof
		if (firmante1Cd ="") then
			firmante1Cd = rsFirmas("CDUSUARIO")
			firmante1Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante1Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante1Rol = CInt(rsFirmas("IDROL"))
			firmante1Sec = rsFirmas("SECUENCIA")
		elseif (firmante2Cd ="") then
			firmante2Cd = rsFirmas("CDUSUARIO")
			firmante2Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante2Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante2Rol = CInt(rsFirmas("IDROL"))
			firmante2Sec = rsFirmas("SECUENCIA")			
		elseif (firmante3Cd ="") then			
			firmante3Cd = rsFirmas("CDUSUARIO")
			firmante3Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante3Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))					
			firmante3Rol = CInt(rsFirmas("IDROL"))
			firmante3Sec = rsFirmas("SECUENCIA")			
		elseif (firmante4Cd ="") then
			firmante4Cd = rsFirmas("CDUSUARIO")
			firmante4Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante4Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante4Rol = CInt(rsFirmas("IDROL"))
			firmante4Sec = rsFirmas("SECUENCIA")
		elseif (firmante5Cd ="") then
			firmante5Cd = rsFirmas("CDUSUARIO")
			firmante5Ds = getUserDescription(rsFirmas("CDUSRROL"))
			firmante5Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			firmante5Rol = CInt(rsFirmas("IDROL"))
			firmante5Sec = rsFirmas("SECUENCIA")
		end if				
		rsFirmas.MoveNext()
	wend	
		
End Function
'--------------------------------------------------------------------------------------------------------------
'Dibuja una tabla en el lugar de las firmas electronicas informando que se bloquearon por alguna determinacion
'	pTitle: Titulo de la tabla
'	pDescripcion: Descripcion de la tabla
Function drawTableBlockSignature(pTitle, pDescripcion)%>
	<table  border="0" cellspacing=0 cellpadding=0 align="center" width="80%">
		<tr>
			<td align="center" colspan="3" height="90px">
				<table width="100%" align="center" border="0" class="reg_header">
					<tr>
						<td align="center" class="reg_header_nav round_border_top">
							<%=pTitle%>
						</td>
					</tr>
					<tr height="90px">
						<td align="center"  class="reg_header_warning">
							<%=pDescripcion%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%End Function
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
idCotizacionElegida = GF_Parametros7("idCotizacionElegida",0,6)
errFirma = GF_PARAMETROS7("errFirma","",6)
accion = GF_PARAMETROS7("accion","",6)

if (errFirma <> "") then Call setError(errFirma)

Call GP_CONFIGURARMOMENTOS

'Leer Cabecera
if idCotizacionElegida > 0 then
	'Se paso como parametro el ID de una cotizacion cargada en el sistema, se debe ir a la base de cotizaciones cargadas.
	strSQL="SELECT * from TBLCTZCABECERA where IDCOTIZACION=" & idCotizacionElegida
	'Response.Write strSQL
	'response.end
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then			
		if (rs("IDPEDIDO") = "0") then
			'Call comprasControlAccesoCM(RES_CD)
			CAB_titulo = "Compra Directa"
		else
			'Call comprasControlAccesoCM(RES_CC)
			CAB_titulo = "Cotizacion Elegida"
		end if
		index = 0
		CAB_importePesos = 0
		CAB_importeDolares = 0
		CAB_observaciones = ""		
		CAB_idCotizacion = rs("IDCOTIZACION")
		CAB_idPedido = rs("IDPEDIDO")			
		Call initHeaderDB(CAB_idPedido)		
		CAB_cdPedido = pct_cdPedido					
		if (CAB_cdPedido = "") then CAB_cdPedido = "Sin Pedido"				
		CAB_idProveedor = rs("IDPROVEEDOR")
		CAB_dsProveedor = getDescripcionProveedor(CAB_idProveedor) 						
		if clng(rs("FECHAENTREGA")) = 0 then
			CAB_fecEntrega = Left(session("MmtoDato"), 8)
		else	
			CAB_fecEntrega = rs("FECHAENTREGA")	
		end if				
		CAB_idObra = rs("IDOBRA")		
		CAB_observaciones = rs("OBSERVACIONES")
		if(CAB_idObra > 0)then
			if(InStr(CAB_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO) > 0)then	
				CAB_observaciones = split(CAB_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO)
				flagDetallePresupuesto = true
			end if
		end if
		CAB_importePesos = CDbl(rs("IMPORTEPESOS"))
		CAB_importeDolares = CDbl(rs("IMPORTEDOLARES"))
		CAB_Moneda = rs("CDMONEDA")
		CAB_IdDivision = rs("IDDIVISION")
		CAB_idContrato = rs("IDCONTRATO")
		CAB_Momento = rs("MOMENTO")	
		if (CLng(CAB_idObra) <> OBRA_GEID) then
		    Call loadDatosObra(CAB_idObra, CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada,CAB_CdResponsable, CAB_DsResponsable)
		    if (CAB_ObraCD = "") then CAB_ObraCD = "Sin Partida"
        else
            CAB_ObraCD = OBRA_GECD
            CAB_ObraDS = OBRA_GEDS
        end if		    
		'Se traen las firmas.
		CAB_estado = rs("ESTADO")
		Call cargarFirmas(idCotizacionElegida)
		'Leer detalles
		Call cargarDetalles(idCotizacionElegida)
		
		Set dicArtUCC = controlarPrecioArticulo(MONEDA_PESO,IT_artID,IT_importePesos,IT_importeDolares,IT_cantidad,CAB_IdDivision, CAB_idCotizacion)
		if (dicArtUCC.count > 0 ) then Call setWarning(PRECIO_DIFIERE_ULTIMO_REGISTRO)
		
		rolUsuario = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
		
	else
		response.redirect "comprasAccesoDenegado.asp"
	end if	
end if	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
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
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE

	var link = "comprasFirmarPIC.asp?idCotizacion=<% =idCotizacionElegida %>&secuencia=";
	var hkey0 = new Hkey('hk0', link + "<%=firmante1Sec %>", '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', link + "<%=firmante2Sec %>"	, '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link + "<%=firmante3Sec %>", '<% =HKEY() %>', 'check_callback()');
	var hkey4 = new Hkey('hk4', link + "<%=firmante4Sec %>", '<% =HKEY() %>', 'check_callback()');
	var hkey5 = new Hkey('hk5', link + "<%=firmante5Sec %>", '<% =HKEY() %>', 'check_callback()');
		
	function check_callback(resp) {			
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}
	
	function volver() {		
		window.close();		
	}
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}		
	
	function bodyOnLoad(){
		var	tb = new Toolbar('toolbar', 8, "images/compras/");
		tb.addButton("Home-16x16.png", "Home", "irHome()");				
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		//Se dibujan los items del detalle		
		tb.draw();		
		hkey0.start();	
		hkey1.start();
		hkey2.start();
		hkey4.start();
		hkey5.start();
	}					
	
	function abrirCTC(id){
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	
	function abrirPCT(id){
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<div id="toolbar"></div><br>	
<table class="reg_header" align="center" width="80%" border="0" >				
	<tr>
		<td colspan="9"><% call showErrors() %></td>
	</tr>
	<tr>
		<td class="reg_header_nav" colspan="9"><% =GF_TRADUCIR("Datos del Pedido") %></td>				
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Ptda. Presup.") %></td>	
		<td colspan="4">
			<%=CAB_ObraCD & " - " & CAB_ObraDS %>
			<input type="hidden" name="CAB_idObra" value="<% =CAB_idObra %>">				
		</td>		
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Pedido") %></td>	
		<td>
			<% if(CAB_idPedido >  0)then %>
				<a><img id="imgPCT" src="images/compras/PCT-16X16.png" style="cursor:pointer" onclick="abrirPCT(<%=CAB_idPedido %>)" title="Abrir Pedido" ></a>&nbsp&nbsp;
			<% end if %>
			<%=CAB_cdPedido%>		
		</td>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Division") %></td>	
		<td colspan="2"><%=getDivisionDS(CAB_idDivision) %></td>
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Proveedor") %></td>
		<td><% =CAB_idProveedor & " - " & CAB_dsProveedor %></td>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Contrato") %></td>
		<td width="3%"></td>		
		<%if(CAB_idContrato > 0)then%>			
		<td>
			<a><img id="imgCTC" src="images/compras/ctc-16x16.png" style="cursor:pointer" onclick="abrirCTC(<%=CAB_idContrato%>)" title="Abrir CTC" ></a>&nbsp&nbsp;
			<%= getCodigoCTC(CAB_idContrato) %>
		</td>		
	    <%end if%>
	</tr>	
	<tr>
		<td class="reg_header_nav" colspan="9"><% =GF_TRADUCIR("Detalle") %></td>
	</tr>	
	<tr>
		<td colspan="8" align="center">
			<table width="90%" id="tblDET">
				<tr>
					<td class="reg_header_nav" width="10%" align="center">	<% =GF_TRADUCIR("Código") %></td>	
					<td class="reg_header_nav" width="40%">	<% =GF_TRADUCIR("Descripción") %></td>	
					<td class="reg_header_nav" width="15%" align="center">	<% =GF_TRADUCIR("Cantidad") %></td>	
					<td class="reg_header_nav" width="10%" align="center">	<% =GF_TRADUCIR("Ptda. Presup.")	%></td>	
					<td class="reg_header_nav" width="25%" align="center" colspan="2">	<% =GF_TRADUCIR("Importe s/IVA") %></td>						
				</tr>
				<%				
				totalImporteCompraUSD = 0
				for index = 0 to ubound(IT_artID) - 1	
					'Si se submitió luego de firmar, no se debe hacer el control de AFEs
					if (accion <> ACCION_CONFIRMAR) then						
						if (incluirArticuloControlAFE(IT_artID(index))) then totalImporteCompraUSD = totalImporteCompraUSD + CDbl(IT_importeDolares(index))
					end if
					IT_Importe = IT_importePesos(index)
					if (CAB_Moneda = MONEDA_DOLAR) then IT_Importe = IT_importeDolares(index)										
					
				%>
					<tr class='<% if (dicArtUCC.Exists(IT_artID(index))) then response.write "reg_header_warning" end if %>'>
						<td align="center"><%=IT_artID(index)%></td>
						<td><%=IT_artDS(index)%></td>	
						<td align="center">	<% = IT_cantidad(index) %></td>	
						<td><% = IT_artBA(index) %> - <% = IT_artBD(index) %></td>						
						<td><% =getSimboloMoneda(CAB_Moneda) %></td>
						<td align="right"><% =GF_EDIT_DECIMALS(IT_Importe,2) %></td>
					</tr>
					<% if(dicArtUCC.Exists(IT_artID(index)))then%>
					<tr>
						<td colspan="6" class="reg_Header_Warning" style='font-weight:  bold;color:#FF0000;'>
							<% =dicArtUCC.Item(IT_artID(index))%>
						</td>
					</tr>
					<%	end if
				next 
				'Si se submitió luego de firmar, no se debe hacer el control de AFEs
				if (accion <> ACCION_CONFIRMAR) then						
					if (necesitaAFE(CAB_idObra, CAB_idPedido,idCotizacionElegida, totalImporteCompraUSD, 0, 0)) then bloqueoPorAFE = true					
				end if				
				%>
					<tr>
						<td colspan="4">&nbsp;</td>	
						<td colspan="2"><hr></td>	
					</tr>
					<tr>
						<td class="reg_header_navdos" colspan="5" align="right"><font size="+1"><b><% =GF_TRADUCIR("Total") %>&nbsp;&nbsp;</b></font></td>							
<%
                        CAB_Importe = CAB_importePesos
					    if (CAB_Moneda = MONEDA_DOLAR) then CAB_Importe = CAB_importeDolares
%>						
						<td align="right"><font size="+1"><b><div id="totalDolaresvisible"><%=getSimboloMoneda(CAB_Moneda) & " " & GF_EDIT_DECIMALS(CAB_importe,2)%></div></b></font></td>							
					</tr>
<%              
    if ((CAB_idPedido = 0) and (CAB_idContrato = 0)) then
        'Es compra directa, si la cantidad del último mes supera el límite maximo de una compra directa se muestra la info.
        mmtoDesde = GF_DTEADD(CAB_Momento,-30,"D")
        Call totalizarComprasDirectasProveedor(CAB_idProveedor, CAB_idDivision, mmtoDesde, CAB_Momento, MONEDAL_DOLAR, cantPICDirectos30, montoPICDirectos30)        
        mmtoDesde = GF_DTEADD(CAB_Momento,-365,"D")
        Call totalizarComprasDirectasProveedor(CAB_idProveedor, CAB_idDivision, mmtoDesde, CAB_Momento, MONEDAL_DOLAR, cantPICDirectos365, montoPICDirectos365)
%>
                    <tr>
						<td class="reg_header_navdos" colspan="4" align="right"><font size="+1"><b><% =GF_TRADUCIR("TOTAL Compras Directas realizadas en los últimos 30 días") %>&nbsp;&nbsp;</b></font></td>							
						<td align="right" colspan="2"><font size="+1"><b><div id="Div2"><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(montoPICDirectos30 - CAB_importeDolares,2) & " (" & cantPICDirectos30 - 1 & " PICs)"%></div></b></font></td>							
					</tr>
					<tr>
						<td class="reg_header_navdos" colspan="4" align="right"><font size="+1"><b><% =GF_TRADUCIR("TOTAL Compras Directas realizadas en los últimos 12 meses") %>&nbsp;&nbsp;</b></font></td>							
						<td align="right" colspan="2"><font size="+1"><b><div id="Div1"><%=getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(montoPICDirectos365 - CAB_importeDolares,2) & " (" & cantPICDirectos365 - 1 & " PICs)"%></div></b></font></td>							
					</tr>
<%        
    end if
%>					
			</table>	
		</td>
	</tr>	

	<tr>
		<td class="reg_header_nav" colspan="9"><% =GF_TRADUCIR("Observaciones") %></td>
	</tr>		
	<tr>
		<td colspan="7" style="height: 100px; vertical-align: top;">
			<% if(flagDetallePresupuesto)then %>
				<div>				
					<%= CAB_observaciones(0) %>
					<%= CAB_observaciones(1) %>
				</div>	
				<div style="color:#FF0000;">
					<b><%= CAB_observaciones(2) %></b>
				</div>
			<% else %>	
				<div>				
					<%= CAB_observaciones %>					
				</div>	
			<% end if %>			
		</td>
	</tr>		
</table>
<form name="frmSel" id="frmSel" method="POST" action="comprasPICFirma.asp?idCotizacionElegida=<% =idCotizacionElegida %>">
<%
if ( bloqueoPorAFE ) then 
	Call drawTableBlockSignature(GF_TRADUCIR("Necesita afe"),GF_TRADUCIR("No se puede autorizar la compra por falta de un AFE que la respalde")) 
else
    flagYaFirmo = false	

%>
	<table align="center" width="80%" border="1" cellspacing=0 cellpadding=0>
		<tr>
			<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Firmas") %></td>
		</tr>
		<tr>
		    <td width="16%"></td>
		    <td width="16%"></td>
		    <td width="16%"></td>
		    <td width="16%"></td>
		    <td width="16%"></td>
		    <td ></td>
		</tr>		
		<tr>
			<td align="center" colspan="2">
				<%	if (firmante1Tx  <> "") then 
				        if (firmante1Cd = session("Usuario")) then flagYaFirmo = true
				%>
					<img src="images/firmas/<% =obtenerFirma(firmante1Cd) %>"><br>
					<% =firmante1Tx %>
				<%	else	
				        if ((session("Usuario") = firmante1Cd) or (rolUsuario = firmante1Rol)) then						
                            flagYaFirmo = true				        
		        %>
							<br><div id="hk0"></div><br>
					<%	else	%>
							<br><br><br>
					<%	end if	
					end if	%>
			</td>
			<td align="center" colspan="2">
				<%	if (firmante2Tx <> "") then 
				        if (firmante2Cd = session("Usuario")) then flagYaFirmo = true
				%>
					<img src="images/firmas/<% =obtenerFirma(firmante2Cd) %>"><br>
					<% =firmante2Tx %>
				<%	else	
				        'response.Write session("Usuario") & "|" & firmante2Cd & "|" & rolUsuario & "|" & firmante2Rol & "|" & isNumeric(firmante2Cd) & "|" & flagBoss
						if (((session("Usuario") = firmante2Cd) or (rolUsuario = firmante2Rol)) and (not flagYaFirmo)) then						
			                flagYaFirmo = true
			    %>
							<br><div id="hk1"></div><br>
					<%	else	%>
							<br><br><br>
					<%	end if	
					end if	%>
			</td>
			<td align="center" colspan="2">
				<%	if (firmante3Tx <> "") then 
				        if (firmante3Cd = session("Usuario")) then flagYaFirmo = true
				%>				
					<img src="images/firmas/<% =obtenerFirma(firmante3Cd) %>"><br>
					<% =firmante3Tx %>
				<%	else	
				        'response.Write "USR Sess:" & session("Usuario") & "|CDUSUARIO:" & firmante3Cd & "|ROL:" & rolUsuario & "|FIRMA ROL:" & firmante3Rol & "|Numerico?:" & isNumeric(firmante2Cd) & "|Jefe:" & flagBoss
						if (((session("Usuario") = firmante3Cd) or (rolUsuario = firmante3Rol))  and (not flagYaFirmo)) then						
				            flagYaFirmo = true		
				%>
							<br><div id="hk2"></div><br>
					<%	else	%>
							<br><br><br>
					<%	end if	
					end if	%>
			</td>
		</tr>
		<tr>
			<td ALIGN="CENTER" colspan="2"><%=firmante1Ds%></td>
			<td ALIGN="CENTER" colspan="2"><%=firmante2Ds%></td>
			<td ALIGN="CENTER" colspan="2"><%=firmante3Ds%></td>										
		</tr>
<%      if ((firmante4Cd <> "") or (firmante5Cd <> "")) then %>		
		<tr>
			<td align="center" colspan="3">
				<%	if (firmante4Tx <> "") then 
				        if (firmante4Cd = session("Usuario")) then flagYaFirmo = true
				%>
					<img src="images/firmas/<% =obtenerFirma(firmante4Cd) %>"><br>
					<% =firmante4Tx %>
				<%	else
						if (((session("Usuario") = firmante4Cd) or (rolUsuario = firmante4Rol)) and (not flagYaFirmo)) then						
				            flagYaFirmo = true		
				%>
							<br><div id="hk4"></div><br>
					<%	else	%>
							<br><br><br>
					<%	end if	
					end if	%>
			</td>
			<td align="center" colspan="3">
				<%	if (firmante5Tx  <> "") then %>
					<img src="images/firmas/<% =obtenerFirma(firmante5Cd) %>"><br>
					<% =firmante5Tx %>
				<%	else	
						if (((session("Usuario") = firmante5Cd) or (rolUsuario = firmante5Rol)) and (not flagYaFirmo)) then						%>
							<br><div id="hk5"></div><br>
					<%	else	%>
							<br><br><br>
					<%	end if	
					end if	%>
			</td>
		</tr>		
		<tr>
			<td ALIGN="CENTER" colspan="3"><%=firmante4Ds%></td>
			<td ALIGN="CENTER" colspan="3"><%=firmante5Ds%></td>
		</tr>
<%      end if %>		
	</table>
<%  end if%>
<input type="hidden" name="errFirma" id="errFirma">
<input type="hidden" name="accion" id="accion" value="<% =ACCION_CONFIRMAR %>">
</form>
</body>
</html>