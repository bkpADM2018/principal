<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosPCT.asp"-->		
<!--#include file="Includes/procedimientosUser.asp"-->

<%
'-----------------------------------------------------------------------------------------------------------
'Esta funcion trae las notas de Aceptacion que estan relacionadas solamente con un Pedido, 
'motrando los que tienen cotizacion = 0 (para la primera tabla) y los de cotizacion > 0(para la segunda tabla).
'Se le pasa como parametro el idPedido
Function getPedidosNDA(pIdPedido)
	Dim strSQL 	
	strSQL = "	SELECT A.IDNDA, A.IDPEDIDO, A.IDCOTIZACION FROM TBLNOTACEPTACION A "	
	strSQL = strSQL & " WHERE A.IDPEDIDO = "&pIdPedido&" and A.IDCOTIZACION = 0 ORDER BY A.IDNDA "
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getPedidosNDA = rs
End function
'-----------------------------------------------------------------------------------------------------------
Function getCotizacionNDA(pIdCotizacion)
	Dim strSQL
	strSQL = " SELECT A.IDNDA, B.IDPEDIDO, A.IDCOTIZACION FROM TBLNOTACEPTACION A "
	strSQL = strSQL & " INNER JOIN TBLCTZCABECERA B ON A.IDCOTIZACION = B.IDCOTIZACION "
	strSQL = strSQL & " WHERE A.IDCOTIZACION = "&pIdCotizacion
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getCotizacionNDA = rs	
End Function
'-----------------------------------------------------------------------------------------------------------
Function getListCotizaciones(pIdPedido,pIdCotizacion)
	Dim strSQL , Mywhere
		if(pIdPedido <> 0)then Mywhere = "IDPEDIDO = "&pIdPedido
		if(pIdCotizacion <> 0)then 
			if(Len(Mywhere) > 0)then Mywhere = Mywhere & " AND "
			Mywhere = Mywhere & " IDCOTIZACION = "&pIdCotizacion
		end if				
	strSQL = "	SELECT IDPEDIDO, IDCOTIZACION,IDPROVEEDOR FROM TBLCTZCABECERA "	
	strSQL = strSQL & " WHERE " & Mywhere & " ORDER BY IDCOTIZACION  "	
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getListCotizaciones = rs
End function
'-----------------------------------------------------------------------------------------------------------
Function getNDA(pIdPedido, pIdCotizacion)
	dim strSQL, rs
	strSQL = "SELECT * FROM tblnotaceptacion WHERE"
	strSQL = strSQL & " IDCOTIZACION = "& pIdCotizacion
	if (pIdPedido <> 0) then strSQL = strSQL & " AND IDPEDIDO = "& pIdPedido
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getNDA = rs
End Function
'*******************************************************************************************************
'********************************   COMIENZO DE LA PAGINA	********************************************
'*******************************************************************************************************
'	Esta pagina es llamada desde ComprasAdministrarCotizaciones.asp y ComprasFichaTab1.asp
'*******************************************************************************************************
Dim colorP, myColor1, myColor2, cont, rsAFE, confirmar, idAFE, rsCompl, myIdDivision, obraPedido
Dim myTitulo,rsPedido,rsCotizacion,idCotizacion,idPedido


idPedido = GF_PARAMETROS7("idPedido",0,6)
idCotizacion = GF_PARAMETROS7("idCotizacion",0,6)
if (idPedido = 0) then	
	' No tiene un pedido asignado, por lo tanto se trae las NDA de una cierta Cotizacion '
	myTitulo = "Compra Directa"	
else
	' Tiene un pedido asignado, por lo tanto se trae todos las NDA de ese pedido '	
	Call initHeaderDB(idPedido)
	' Cargo las variables del PCT (Titulo, Codigo)
	myTitulo = pct_tituloPedido 
	Set rs = getPedidosNDA(idPedido)
end if


%>
<html>
<head>
<title><%=GF_TRADUCIR("Administrar NDA")%></title>
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>	
<script type="text/javascript">
	var ch= new channel();
	var popUpNDA;
		
	function agregarNDA(pIdPedido,pIdCotizacion,pIdProv){		
		popUpNDA = winPopUp('Iframe', "comprasNDAPopUp.asp?idPedido=" + pIdPedido + "&idcotizacion="+pIdCotizacion+"&accion=<%=ACCION_LEER_NDA%>&IdProveedor="+pIdProv , "450", "300", '<%=GF_Traducir("Detalle Nota de Aceptacion")%>', 'reloadPage()');
	}		
	
	function reloadPage(){
		window.location.reload();
	}	
	
	function actualizarNDA(idPedido,idcotizacion,accion,idNDA){
		ch.bind("comprasNDAAjax.asp?idPedido="+idPedido+"&idCotizacion="+idcotizacion+"&accion="+accion+"&IdNDA="+idNDA,"envioNDA_callBack("+idPedido+","+idcotizacion+","+idNDA+")");
		ch.send();			
	}	
	function envioNDA_callBack(idPed,idCot,idNDA) {		
		ch.bind("comprasEnvioNDAMail.asp?idPedido=" + idPed +"&idCotizacion="+ idCot +"&IdNDA="+idNDA,"");
		ch.send();		
	}	
	function cerrarPopUpNDA()
	{
		popUpNDA.hide();
	}
	
</script>
</head>
<body  >	
<form id="frmSel" name="frmSel" action="comprasListaNDA.asp" method="POST"  >		
	<table>
		
		<tr>
			<div id="avisoNDA" align="center" class="TDBAJAS"></div>			
		</tr>
		<% if(idPedido > 0)then 
			' En caso de que venga con un Pedido,muestra las NDA del Pedido en caso de que tenga.
			' En caso de que no tenga un Pedido asociado, no va a mostrar esta tabla
		%>
		<tr>
			<table id="TblPCT" align="center" class="reg_Header" width="80%" >
				<tr><td colspan="4"><h4><% =GF_TRADUCIR("Pedido: ")&pct_cdPedido &" - "&myTitulo%></h4></td></tr>	
				<tr class="reg_Header_nav">
					<td colspan="3"><%= "PCT"&" - "&pct_tituloPedido %></td>
					<td align="center" width="10%" class="reg_Header"><img src="images/compras/add-16x16.png"title="Nuevo NDA"  style="cursor:pointer" onclick="agregarNDA(<%=idPedido%>,0,0)"></td>
				</tr>
				<%	 while not rs.eof  %>
						<tr class="reg_Header_navdos">
							<td width="10%" align="center"><img src="images/compras/NDA-16x16.png"></td>
							<td align="left" width="80%"><% =GF_TRADUCIR("NDA")& " - " & rs("IDNDA") %></td>
							<td align="center" width="10%" colspan="2"><a href="comprasNotaDeAceptacionPrint.asp?idPedido=<% =rs("IDPEDIDO") %>&idcotizacion=0&idNDA=<%=rs("IDNDA")%>" target="_blank"><img style="cursor:pointer" src="images/compras/printer-16x16.png" title="Imprimir NDA"></a></td>			 				
						</tr>
				<%	  rs.MoveNext()
					 wend 
			 %>			
			</table>
		</tr>	
		<% end if %>	
		<br></br>		
		<%	' Listamos todas las cotizaciones que tiene el pedido, si viene si pedido solo muestra una sola Cotizacion
		    Set rsPIC = getListCotizaciones(idPedido,idCotizacion)	
			if(not rsPIC.eof)then
		%>
		<tr>
			<table id="TblPIC" align="center" class="reg_Header" width="80%" >
			<% ' Si viene con un Pedido asociado , va a mostrar todas las Cotizaciones que lo conforman junto con sus NDA  
			   ' En cambio si no tiene un Pedido y solo trae la Cotizacion va a mostrar solo los NDA de esa cotizacion %>				
			<%  while not rsPIC.eof  %>
				<tr  class="reg_Header_nav" >
					<td colspan="3"><%= "PIC"&" - "&rsPIC("IDCOTIZACION") %></td>
					<td align="center" width="10%" class="reg_Header"><img style="cursor:pointer" title="Nuevo NDA" src="images/compras/add-16x16.png" onclick="agregarNDA(<%=idPedido%>,<%=rsPIC("IDCOTIZACION")%>,<%=rsPIC("IDPROVEEDOR")%>)"></td>
				</tr>									
			<% ' Se busca las NDA del PIC en caso de que tenga
					Set rsNDA = getNDA(idPedido, rsPIC("IDCOTIZACION"))
				 	while not rsNDA.eof %>							 
					 <tr class="reg_Header_navdos" >
						 <td width="10%" align="center"><img src="images/compras/NDA-16x16.png"></td>
						 <td width="70%" align="left"><% =GF_TRADUCIR("NDA")& " - " & rsNDA("IDNDA") %></td>
				 		 <td width="10%" align="center"><a href="comprasNotaDeAceptacionPrint.asp?idPedido=<% =rsNDA("IDPEDIDO") %>&idcotizacion=<%=rsNDA("IDCOTIZACION")%>&idNDA=<%=rsNDA("IDNDA")%>" target="_blank"><img style="cursor:pointer" src="images/compras/printer-16x16.png" title="Imprimir NDA"></a></td>
						 <td align="center"><a href="javascript:actualizarNDA(<% =rsNDA("IDPEDIDO") %>,'<%=rsNDA("IDCOTIZACION")%>','<% =ACCION_ACTUALIZAR_NDA%>','<%=rsNDA("IDNDA")%>')"><img id="iconNDA_C_<%=rsNDA("IDNDA")%>" style="cursor:pointer" src="images/compras/mail_sent-16x16.png" title="Enviar NDA"></a></td>
					 </tr>							
				<%   rsNDA.MoveNext()
					 wend			 						
				rsPIC.MoveNext()
			    wend     %>									
			</table>			
		</tr>	
		<% end if %>	
	</table>	
</form>
</body>
</html>