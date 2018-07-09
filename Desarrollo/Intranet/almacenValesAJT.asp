<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call controlAccesoAL("")

Const ITEM_ID = "item"
Const ITEM_AJUSTE = "cumplido"
Const ITEM_CANTIDAD = "amount"
Const ITEM_SALDO = "saldo"
	
Function hayPMReferencia() 
	dim strSQL, rs, km, kc, tmp, pos	
	hayPMReferencia = false		
	if not (idPMReferencia = 0) then
		strSQL="select * from TBLVALESCABECERA where PARTIDAPENDIENTE=" & idPMReferencia & " AND ESTADO=" & ESTADO_ACTIVO & " AND CDVALE IN ('" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_TRANSFERENCIA & "')"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			if rs("CDVALE") = CODIGO_VS_PRESTAMO then esPrestamo = true
			hayPMReferencia = true
		else
			setError(PM_REFERENCIA_NO_EXISTE_NO_TIPO)
		end if
	end if
End Function
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idVale, index, esModificable, controlOK, submitPage, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim rsObras, rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myvaleSalidaComment
dim minCantPro, nrmName, fromTC, aux, auxAju
dim rsAlmacenes, devueltos, esPrestamo
dim title1, title2, title3, mainTitle, lastChar
dim idPMReferencia, idPMReferenciaHDDN, myJSClose, flagGrabarVale
esPrestamo = false
idPMReferencia = GF_PARAMETROS7("pmReferencia",0,6)
fromTC = GF_PARAMETROS7("TC",0,6)
VS_cdVale = GF_PARAMETROS7("cdVale","",6)
idPMReferenciaHDDN = GF_PARAMETROS7("pmReferenciaHDDN",0,6)
accion = GF_PARAMETROS7("accion","",6)
VS_comentario = GF_PARAMETROS7("comentario","",6)
call GP_ConfigurarMomentos()


idVale = 0
controlOK = false
flagGrabarVale = false
estaPMReferencia = hayPMReferencia()
if (not isFormSubmit()) then
	'No submitio la pagina, primera vez que entra
	if estaPMReferencia then	
		call initHeaderPMDB(idPMReferencia)
		call PM2VS		
	else
		vs_idObra = 0
		vs_idAlmacen = 0
	end if
else	
	'Submitio la pagina		
	if idPMReferencia <> 0 then
		call initHeaderPMDB(idPMReferencia)
		call PM2VS
		call initArticulosDB(idPMReferencia)
		VS_CantArticulos = PM_CantArticulos
		VS_ArticuloActual = PM_ArticuloActual
		VS_FechaSolicitud = GF_PARAMETROS7("issuedate", "", 6)	
		if (VS_FechaSolicitud = "") then VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		

		'Controlar el Vale
		controlOK = controlarVale(idPMReferencia)
		if ((accion = ACCION_GRABAR) and (controlOK)) then
			VS_ArticuloActual = 0
			flagGrabarVale = true
			call grabarHeaderVale(idVale, idPMReferencia)
			call grabarComentarioVale(idVale, VS_comentario)
			while (readNextArticuloVale(idVale))
				if (VS_saldo > 0) then
					call grabarValeDetalle(idVale, idPMReferencia)
				end if
			wend
			Call grabarPreciosVigentesPorArticulo(idVale)
			if (fromTC <> 1) then		
				myJSClose = "location.href= 'almacenAjustes.asp'"
			else	
				myJSClose = "cerrar();"
			end if								
		end if
	else
		call setError (PM_REQUERIDO) 	
	end if
end if
%>
<html>
<head>
<title><%=GF_TRADUCIR("Almacen - Vales de Ajuste")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
</style>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	
	<% if flagGrabarVale then %>
		if (confirm("Desea realizar la impresion del vale?")) {
			window.open('almacenValePedidoPrint.asp?idVale=<% =idVale %>','Imprimir Vale');
		}		
	<% end if %>
	
	//Constantes - Nombre de Campo
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";	
	var ITEM_DIV = "itemDiv";
	var ITEM_DESC = "articuloItem";
	var ITEM_AMOUNT = "amount";
	var ITEM_AMOUNT_TEXT = "amountText";
	var ITEM_SALDO = "saldo";
	var ITEM_CUMPLIDO = "cumplido";
	var ITEM_AMOUNT_UNIT = "abreviatura";	
	var ITEM_CUMPLIDO_TEXT = "cumplidoText";
	var ITEM_SALDO_TEXT = "saldoText";
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");			
	var tb;		
	var lastArticulos = 0;				
	
	
	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}		
	
	function cerrar() {	
		<%	if (fromTC = 1) then %>
				var refpopupAJT;
				refpopupAJT = startIWin('popupAJT');
				refpopupAJT.hide();
		<%	else 	%>
				window.close();			
		<%	end if	%>
	}
	function irAjustes() {
		location.href = "almacenAjustes.asp";
	}	
	function bodyOnLoad() {
		<%=myJSClose%>
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>',0)");
		tb.addButtonCONFIRM("Controlar", "submitInfo('<% =ACCION_CONTROLAR %>',1)");
		<% if (fromTC <> 1) then %> 
			tb.addButton("Setting_folder-16x16.png", "Ajustes", "irAjustes()");	
		<% else %> 
			tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		<% end if %> 
		tb.draw();		
		<%	
		index = 0
		if ((estaPMReferencia)) then
			if (idPMReferenciaHDDN <> idPMReferencia) then
				PM_hayCabecera = true
				call initHeaderPMDB(idPMReferencia)
				call initArticulosDB(idPMReferencia)
				while (readNextArticuloDB())					
					PM2VS_DET
					auxAju = getTotalArticuloxVale(idPMReferencia, vs_idarticulo, CODIGO_VS_AJUSTE_TRANSFERENCIA)
					if (CLng(auxAju) <> 0) then VS_cantidad = CDbl(VS_cantidad) - CDbl(auxAju)					
					devueltos = getCantidadRecibida(idPMReferencia, VS_idArticulo)						
					
					VS_cumplido = clng(VS_cantidad) - clng(VS_saldo)
					VS_saldo = VS_cumplido - clng(devueltos)
					%>
					agregarLineaArticulo();
					fillArticulo(<% =index %>, '<% =vs_idArticulo %>', '<% =vs_dsArticulo %>', <% =VS_cumplido %>, <% =devueltos %>, <% =VS_saldo %>, '<% =vs_abreviaturaUnidad %>');					
					<%
					index=index+1
				wend
			else	
				while (readNextArticuloValeParams())
		 			%>
					agregarLineaArticulo();
					fillArticulo(<% =index %>, '<% =VS_idArticulo %>', '<% =VS_dsArticulo %>', <% =VS_cantidad %>, <% =VS_cumplido %>, <% =VS_saldo %>, '<% =vs_abreviaturaUnidad %>');					
					<%
					index=index+1
				wend
			end if	
		end if
	%>		
		pngfix();
		resaltarArticulosConErrores();	
	}
	
	function resaltarArticulosConErrores(){
		//resalta con otro color articulos con errores, que consigue del array arrArticulosConErrores
		var tblArticulos = document.getElementById("tblArticulos");
		<%
		dim iArticulos
		For iArticulos = 0 to ubound(arrArticulosConErrores)%>
	      for (i=0; i< <%=index%>;i++){
			if (document.getElementById('item' + i).value == '<%=arrArticulosConErrores(iArticulos)%>') tblArticulos.rows[i+1].className = 'reg_Header_Error';		      
		  }
	    <%next%>
	}
	function agregarLineaArticulo() {		
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var index;
		index = 2;
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCantidad = rArticulo.insertCell(2);
		var cCumplido = rArticulo.insertCell(3);		
		var cSaldo = rArticulo.insertCell(4);	
		
		
		//CODIGO
		var iCodigo = document.createElement('input');
		iCodigo.type = "hidden";
		iCodigo.id = ITEM_ID + lastArticulos;
		iCodigo.name = ITEM_ID + lastArticulos;
		iCodigo.size= 7;
		iCodigo.maxLength = 5;				
		cCodigo.appendChild(iCodigo);			
		var dCodigo = document.createElement('div');
		dCodigo.className = "labelStyle";
		dCodigo.id = ITEM_DIV + lastArticulos;		
		cCodigo.appendChild(dCodigo);
		
		//DESCRIPCION
		var iDescripcion = document.createElement('div');		
		iDescripcion.id = ITEM_DESC + lastArticulos;				
		cDescripcion.appendChild(iDescripcion);		
		
		
		//CANTIDAD	
		cCantidad.align = 'right';
		var iCantidad = document.createElement('input');	
		var dCantidadUnidad = document.createElement('span');
		var dCantidad = document.createElement('span');		
		iCantidad.type = "hidden";
		<% if idVale > 0 or estaPMReferencia then %>	
			iCantidad.type = "hidden";
			dCantidadUnidad.style.display = 'none';
		<% end if %>	
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		iCantidad.size= 4;
		iCantidad.align = 'center';
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");			
		} else {
			iCantidad['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}
		iCantidad.id = ITEM_AMOUNT + lastArticulos;		
		cCantidad.appendChild(iCantidad);
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		dCantidadUnidad.style.textAlign = "right";
		iCantidad.style.textAlign = "right";
		cCantidad.appendChild(dCantidadUnidad);		
		
		//CUMPLIDO TEXT	
		dCantidad.id = ITEM_AMOUNT_TEXT + lastArticulos;	
		cCantidad.appendChild(dCantidad);
		
		//CUMPLIDO
		cCumplido.align = 'right';
		var iCumplido = document.createElement('input');
		var dCumplido = document.createElement('span');
		iCumplido.type = "hidden";
		iCumplido.name = ITEM_CUMPLIDO + lastArticulos;
		iCumplido.size = 4;
		if (isFirefox) {
			iCumplido.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");			
		} else {
			iCumplido['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}
		iCumplido.id = ITEM_CUMPLIDO + lastArticulos;
		<% if VS_cdVale = CODIGO_VS_ENTRADA or VS_cdVale = CODIGO_PM then %>
			cCumplido.style.display = 'none';
			dCumplido.style.display = 'none';
		<% end if %>
		cCumplido.appendChild(iCumplido);	
		
		//CUMPLIDO TEXT	
		dCumplido.id = ITEM_CUMPLIDO_TEXT + lastArticulos;	
		iCumplido.style.textAlign = "right";
		cCumplido.appendChild(dCumplido);
		
		//SALDO
		cSaldo.align = 'right';
		var iSaldo = document.createElement('input');	
		var dSaldo = document.createElement('span');	
		iSaldo.name = ITEM_SALDO + lastArticulos;
		iSaldo.size= 4;
		if (isFirefox) {
			iSaldo.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");					
		} else {
			iSaldo['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}		
		iSaldo.id = ITEM_SALDO + lastArticulos;		
		cSaldo.appendChild(iSaldo);
		<%if (idVale > 0) or VS_cdVale = CODIGO_VS_ENTRADA or VS_cdVale = CODIGO_PM then %>
			iSaldo.type = "hidden";
			dSaldo.style.display = 'none';
		<% end if %>

		//SALDO TEXT
		dSaldo.id = ITEM_SALDO_TEXT + lastArticulos;
		iSaldo.style.textAlign = "right";
		cSaldo.appendChild(dSaldo);
		
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;		
	}
	
	function fillArticulo(linea, id, desc, cantidad, cumplido, saldo, unit) {
		document.getElementById(ITEM_DIV + linea).innerHTML = id;
		document.getElementById(ITEM_ID + linea).value = id;			
			
		document.getElementById(ITEM_DESC + linea).innerHTML = desc;					

		if (document.getElementById(ITEM_AMOUNT + linea).type != "text"){
			document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = cantidad + " " + unit;
		}
		else{
			document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = "&nbsp;" + unit;
		}

		if (document.getElementById(ITEM_SALDO + linea).type != "text"){
			document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = saldo + " " + unit;
		}	
		else{
			document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = "&nbsp;" + unit;
		}

		if (document.getElementById(ITEM_CUMPLIDO + linea).type != "text"){
			document.getElementById(ITEM_CUMPLIDO_TEXT + linea).innerHTML = cumplido + " " + unit;
		}	
		else{
			document.getElementById(ITEM_CUMPLIDO_TEXT + linea).innerHTML = "&nbsp;" + unit;
		}
		document.getElementById(ITEM_AMOUNT + linea).value = cantidad;
		document.getElementById(ITEM_CUMPLIDO + linea).value = cumplido;
		document.getElementById(ITEM_SALDO + linea).value = saldo;
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">	
<div id="toolbar"></div>
<br>		
<%
if (fromTC = 1) then
	submitPage = "almacenValesAJT.asp"
else
	submitPage = "almacenValesTitulo.asp?cdVAle=" & CODIGO_VS_AJUSTE_TRANSFERENCIA
end if	
%>
<form id="frmSel" name="frmSel" action="<% =submitPage %>" method="POST">	
	<table class="reg_Header" align="center" width="70%" border="0">				
		<tr>
			<td colspan="6">
				<%call showErrors()%>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" colspan="1" align="center"><font class="big"><% =ucase(VS_cdVale) %></font></td>
			<td align="right" class="reg_Header_nav" colspan="5">
				<% 
				Response.write GF_TRADUCIR("Nº Pedido de Materiales de Referencia:") 
				%>
				<input id="pmReferenciaHDDN" type="hidden" name="pmReferenciaHDDN" value="<% =idPMReferencia %>">
				<input id="pmReferencia" type="text" size="5" name="pmReferencia" value="<% =idPMReferencia %>">
			</td>
		</tr>		
        <tr>
			<td class="reg_Header_navdos"><%= GF_TRADUCIR("Part. Pres.") %></td>
			<td colspan="2">
				<%
					if vs_idObra <> 0 then
						Set rsObra = obtenerListaObras(vs_idObra, "", "","","") 
						if (not rsObra.eof) then 
							response.write rsObra("CDOBRA") & " - " & rsObra("DSOBRA")
						end if	
						%>
						<input type="hidden" name="idObra" id="idObra" value="<% =vs_idObra %>">					
						&nbsp;&nbsp;&nbsp;<span id="secBudgetDiv"></span>
						<%
					else
						%>
						<input type="hidden" name="idSector" id="idSector" value="<% =vs_idSector %>">					
						&nbsp;&nbsp;&nbsp;<% =vs_idSector %>
						<%
					end if						
					%>
			</td>
			
		</tr>
		<tr>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Solicitante") %></td>
			<td colspan="2">
				<%
				response.write VS_dsSolicitante
				%>
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =VS_cdSolicitante %>"/>
			</td>
			<td width="15%"  class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
			<td colspan="2">
				<% 
				Set rsAlmacenes = obtenerListaAlmacenes(vs_idAlmacen) 
				if (not rsAlmacenes.eof) then 
				    response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
				end if
				%>
				<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =vs_idAlmacen %>">
			</td>
		</tr>	
		<tr>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha Ajuste") %></td>
			<td colspan="5" align="left">
				<table width=100% cellspacing=0 cellpadding=0>
					<tr>
						<td align=left>
							<div id="issuedateDiv"><% =VS_FechaSolicitud %></div>
							<input type="hidden" id="issuedate" name="issuedate" value="<% =VS_FechaSolicitud %>">
						</td>						
					</tr>
				</table>
			</td>			
		</tr>		
		<tr>
		<%	if (not esPrestamo) and hayPMReferencia then %>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen Destino") %></td>
			<td colspan="5">
				<%
					Set rsAlmacenes = obtenerListaAlmacenes(VS_idAlmacenDest) 
					if (not rsAlmacenes.eof) then 
						response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
					end if
				%>
				<input type="hidden" id="idAlmacenDest" name="idAlmacenDest" value="<% =VS_idAlmacenDest %>">
			</td>				
		<%	end if%>
		</tr>		
		<% if ((estaPMReferencia)) then %>
			<tr>
				<td class="reg_Header_nav" colspan="6"><% =GF_TRADUCIR("Comentario") %></td>
			</tr>
			<tr>
				<% if idVale <> 0 then %>
					<td colspan="6"><%=getComentarioVale(idVale)%></td>
				<% else	%>
					<td colspan="6" align=center><textarea name="comentario" id="comentario" cols="100"><%=VS_comentario%></textarea>
				<% end if %>
				</td>
			</tr>
			<tr>
				<td class="reg_Header_nav" colspan="6"><% =GF_TRADUCIR("Detalle") %></td>
			</tr>
					
			<tr>
				<td colspan="6">
					<table class="reg_Header" width="100%" id="tblArticulos">
						<tr class="reg_Header_nav">
							<td align="center" width="10%"><% =GF_TRADUCIR("Codigo") %></td>
							<td align="center" width="50%"><% =GF_TRADUCIR("Descripcion") %></td>
							<td align="center" width="10%"><% =GF_TRADUCIR("Transferidos") %></td>
							<td align="center" width="10%"><% =GF_TRADUCIR("Recibidos") %></td>
							<td align="center" width="10%"><% =GF_TRADUCIR("Ajuste") %></td>
						</tr>					
					</table>
				</td>
			</tr>
		<% end if %>
	</table>
	<input type="hidden" id="accion" name="accion" value="">
	<input type="hidden" id="TC" name="TC" value="<% =fromTC %>">
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="<% =index%>">	
	<input type="hidden" id="cdVale" name="cdVale" value="<% =VS_cdVale %>">
	<input type="hidden" name="resp" id="resp" value="MAYBE">			
	<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =VS_cdSolicitante %>"/>
</form>
</body>
</html>
<%
'---------------------------------------------------------------------------------------------
sub PM2VS()
	'VS = PM
	VS_FechaSolicitud = PM_FechaSolicitud
	VS_FechaRequerido = PM_FechaRequerido
	VS_cdSolicitante = PM_cdSolicitante
	VS_dsSolicitante = PM_dsSolicitante
	VS_idPedido = PM_idPedido
	VS_idAlmacen = PM_idAlmacen
	VS_idAlmacenDest = PM_idAlmacenDest
	VS_idObra = PM_idObra
	VS_idBudgetArea = PM_idBudgetArea
	VS_idBudgetDetalle = PM_idBudgetDetalle
	VS_idSector = PM_idSector
	VS_usuario = PM_usuario
	VS_momento = PM_momento
	VS_hayCabecera = PM_hayCabecera
end sub
'---------------------------------------------------------------------------------------------
sub PM2VS_DET()
	VS_idArticulo = PM_idArticulo
	VS_dsArticulo = PM_dsArticulo	
	VS_idUnidad = PM_idUnidad
	VS_abreviaturaUnidad = PM_abreviaturaUnidad
	VS_cantidad = PM_cantidad
	VS_saldo = PM_saldo
end sub
%>