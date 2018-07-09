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
'--------------------------------------------------------------------
Function hayPMReferencia()
	dim strSQL, rs, km, kc, tmp, pos	
	hayPMReferencia = false
	if not (idPMReferencia = 0) then
		strSQL="select * from TBLPMCABECERA where IDPEDIDO=" & idPMReferencia
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			hayPMReferencia = true
		else
			setError(PM_REFERENCIA_NO_EXISTE)
			hayPMReferencia = false
		end if
	end if
End Function
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idVale, index, esModificable, controlOK, submitPage, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim rsObras, rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myvaleSalidaComment
dim minCantPro, nrmName, fromTC, aux
dim rsAlmacenes, devueltos,rsBudget, flagHeader
dim title1, title2, title3, mainTitle, lastChar
dim idPMReferencia, idPMReferenciaHDDN, myJSClose, flagGrabarVale
dim myIdAlmacen, auxSaldo

call GP_ConfigurarMomentos()

myIdAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
PM_idAlmacen = myIdAlmacen
idVale = GF_PARAMETROS7("idVale",0,6)
idPMReferencia = GF_PARAMETROS7("pmReferencia",0,6)
VS_cdVale = GF_PARAMETROS7("cdVale","",6)
fromTC = GF_PARAMETROS7("TC",0,6)
idPMReferenciaHDDN = GF_PARAMETROS7("pmReferenciaHDDN",0,6)
VS_comentario = GF_PARAMETROS7("comentario","",6)
accion = GF_PARAMETROS7("accion","",6)
resp = GF_PARAMETROS7("resp","",6)
myJSClose = ""
controlOK = false
estaPMReferencia = hayPMReferencia()

'Se controla si se puede acceder al pedido.
if (not checkControlPM(idPMReferencia)) then response.redirect "comprasAccesoDenegado.asp"
if (not isFormSubmit()) then	
	'No submitio la pagina
	if (idVale > 0) then 
		'Entra a ver un Vale preexistente.
		call initHeaderVale(idVale)
		VS_cdSolicitante = session("Usuario")
		VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
		if ((idVale = 0) and (myIdAlmacen > 0)) then	VS_idAlmacen = myIdAlmacen 		
		call initArticulosVale(idVale)
	else
		'Primera vez que entra
		'Cargo un PM, leer info desde alli
		call initHeaderPMDB(idPMReferencia)
		call PM2VS()
		VS_cdSolicitante = session("Usuario")
		VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
		call initArticulosDB(idPMReferencia)
	end if
else	
	'Submitio la pagina	
	call initHeaderVale(idVale)
	VS_cdSolicitante = session("Usuario")
	VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
	call initArticulosVale(idVale)	
	VS_FechaSolicitud = GF_PARAMETROS7("issuedate", "", 6)	
	if (VS_FechaSolicitud = "") then VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		

	'Controlar el Vale
	controlOK = controlarVale(idVale)
	if ((accion = ACCION_GRABAR) and (controlOK)) then		
	    flagHeader = false
		while ((readNextArticuloVale(idVale)))
				if VS_cantidad <> VS_cumplido then
				    'Si se graba el primer artículo, entonces grabo la cabecera.
				    if (not flagHeader) then
				        Call grabarHeaderVale(idVale, idPMReferencia)
		                Call grabarComentarioVale(idVale, VS_comentario)
		                flagHeader = true
		            end if
					Call actualizarPMDetalleAju(idPMReferencia, VS_idArticulo, VS_cantidad, VS_cumplido, VS_Saldo)
					VS_saldo = VS_cantidad - VS_cumplido
					call grabarValeDetalle(idVale, idPMReferencia)
				end if
		wend
		myJSClose = "cerrar();"
	end if
end if
%>
<html>
<head>
<title><%=GF_TRADUCIR("Almacen - Vales")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
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
<script type="text/javascript" src="scripts/uploadManager.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	
	<% if flagGrabarVale then %>
		if (confirm("Desea realizar la impresion del vale?")) {
			window.open('almacenValePedidoPrint.asp?idVale=<%=idVale%>','Imprimir Vale');
		}
		
	<% end if %>
	
	//Constantes - Nombre de Campo
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";
	var ITEM_DESC = "itemDesc";
	var ITEM_DIV = "itemDiv";
	var ITEM_CANTIDAD = "amount";
	var ITEM_CANTIDAD_UNIT = "amount_unit";
	var ITEM_CANTIDAD_TEXT = "amount_text";
	var ITEM_SALDO = "saldo";
	var ITEM_ENTREGADO_UNIT = "saldo_unit";
	var ITEM_ENTREGADO_TEXT = "saldo_text";
	var ITEM_NUEVA = "cumplido";	
	var ITEM_NUEVA_UNIT = "cumplido_unit";	


	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var lastArticulos = 0;		
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ch = new channel();		
	function aceptarvaleSalida(id) {
		document.getElementById("resp").value = "OK";
		document.getElementById("frmSel").submit();
	}
	
	function agregarLineaArticulo() {		
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var index;
		index = 1;
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCantidad = rArticulo.insertCell(2);
		var cSaldo = rArticulo.insertCell(3);		
		var cNueva = rArticulo.insertCell(4);	
	

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

		//Cantidad
		var sCantidad = document.createElement('span');
		sCantidad.id = ITEM_CANTIDAD_TEXT + lastArticulos;	
		sCantidad.align = 'right';
		var sCantidadUnidad = document.createElement('span');
		sCantidadUnidad.id = ITEM_CANTIDAD_UNIT + lastArticulos;	
		sCantidadUnidad.align = 'right';
		var iCantidad = document.createElement('input');		
		iCantidad.type = "hidden";
		iCantidad.id = ITEM_CANTIDAD + lastArticulos;
		iCantidad.name = ITEM_CANTIDAD + lastArticulos;
		cCantidad.align = 'right';
		cCantidad.appendChild(sCantidad);		
		cCantidad.appendChild(sCantidadUnidad);
		cCantidad.appendChild(iCantidad);				
		
		//Saldo
		var sSaldo = document.createElement('span');
		sSaldo.id = ITEM_ENTREGADO_TEXT + lastArticulos;	
		sSaldo.align = 'right';
		var sSaldoUnidad = document.createElement('span');
		sSaldoUnidad.id = ITEM_ENTREGADO_UNIT + lastArticulos;	
		sSaldoUnidad.align = 'right';
		var iSaldo = document.createElement('input');		
		iSaldo.type = "hidden";
		iSaldo.id = ITEM_SALDO + lastArticulos;
		iSaldo.name = ITEM_SALDO + lastArticulos;
		cSaldo.align = 'right';
		cSaldo.appendChild(sSaldo);		
		cSaldo.appendChild(sSaldoUnidad);
		cSaldo.appendChild(iSaldo);				

		//Nueva cantidad
		var iNueva = document.createElement('input');
		iNueva.name = ITEM_NUEVA + lastArticulos;
		iNueva.id = ITEM_NUEVA + lastArticulos;
		iNueva.size = 4;
		iNueva.style.align = 'right';
		if (isFirefox) {
			iNueva.setAttribute('onkeypress', "return controlDatos(this, event, 'N')");
			//iNueva.setAttribute('onblur', "return controlCampo(this, 'N')");
		} else {
			iNueva['onkeypress'] = new Function("return controlDatos(this, event, 'N')");
			//iNueva['onblur'] = new Function("return controlCampo(this, 'N')");
		}
		var sNuevaUnidad = document.createElement('span');
		sNuevaUnidad.id = ITEM_NUEVA_UNIT + lastArticulos;	
		
		cNueva.align = 'right';
		cNueva.appendChild(iNueva);	
		cNueva.appendChild(sNuevaUnidad);	
		
		var ms;
<%		if (idVale = 0) then 				%>
			ms = new MagicSearch("", ITEM_DESC + lastArticulos, 50, 4, "comprasStreamElementos.asp?tipo=articulos");
			ms.setToken(";");
			ms.onBlur = "seleccionarArticulo(" + lastArticulos + ")";
<%		end if	%>
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;
		return ms;
	}

	function seleccionarArticulo(linea, vss) {
		var desc = "";
		if (vss){
			if (typeof(vss) != "boolean") desc = vss.getSelectedItem();
		}			
		if (desc.indexOf('|') != -1) {					
			var arr = desc.split('|');
			document.getElementById(ITEM_ID + linea).value = arr[0];
			document.getElementById(ITEM_DIV + linea).innerHTML = arr[0];			
			var arr2 = arr[1].split('[');	
			document.getElementById(ITEM_PEDIDO_AJU + linea).innerHTML = arr2[1].replace(/]/,"");
			document.getElementById(ITEM_ENTREGADO_AJU + linea).innerHTML = arr2[1].replace(/]/,"");
			vss.setValue(arr2[0]);			
		} else {
			if (desc == "") {
				document.getElementById(ITEM_ID + linea).value = "";
				document.getElementById(ITEM_DIV + linea).innerHTML = "";
				document.getElementById(ITEM_PEDIDO_AJU + linea).innerHTML = "";
				document.getElementById(ITEM_ENTREGADO_AJU + linea).innerHTML = "";
			}
		}
	}
	function seleccionarSolicitante(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdSolicitante").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdSolicitante").value = "";							
		}		
	}
	
	function fillArticulo(vss, linea, id, desc, cantidad, saldoActual, nueva, unit) {	
<%		if (esModificable) then 				%>
			vss.setValue(id + "-" + desc + "[" + unit + "]");
			seleccionarArticulo(linea, vss);		
<%		else				 				%>		
			document.getElementById(ITEM_DIV + linea).innerHTML = id;
			document.getElementById(ITEM_ID + linea).value = id;
			<%if not estaPMReferencia then%>
				vss.setValue(desc);
			<% else %>				
				document.getElementById(ITEM_DESC + linea).innerHTML = desc;		
			<% end if %>	

			document.getElementById(ITEM_DIV + linea).innerHTML = id;
			document.getElementById(ITEM_ID + linea).value = id;
			document.getElementById(ITEM_DESC + linea).innerHTML = desc;		
			document.getElementById(ITEM_CANTIDAD + linea).value = cantidad;
			document.getElementById(ITEM_CANTIDAD_TEXT + linea).innerHTML = cantidad;
			document.getElementById(ITEM_CANTIDAD_UNIT + linea).innerHTML = "&nbsp;" + unit;			
			document.getElementById(ITEM_SALDO + linea).value = saldoActual;
			document.getElementById(ITEM_ENTREGADO_TEXT + linea).innerHTML = cantidad - saldoActual;
			document.getElementById(ITEM_ENTREGADO_UNIT + linea).innerHTML = "&nbsp;" + unit;
			document.getElementById(ITEM_NUEVA + linea).value = nueva;			
			document.getElementById(ITEM_NUEVA_UNIT + linea).innerHTML = "&nbsp;" + unit;
<%		end if				 			%>
	}

	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}	
	function canSubmit(acc, btn) {		
			submitInfo(acc);		
	}
	function cerrar() {	
		<%	if fromTC <> 1 then %> 
				window.close();
		<%	else	%>
				var refPopUpArt;
				refPopUpArt = startIWin('popupAJP');
				refPopUpArt.hide();		
		<%	end if %>
	}	
	function irHome() {
		location.href = "almacenIndex.asp";
	}	
	function irAjustes() {
		location.href = "almacenAjustes.asp";
	}	
	function bodyOnLoad() {	
		<% =myJSClose%>
		var myMS;
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		<% if fromTC <> 1 then %> 
			tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% end if %> 
		<%	if (idVale = 0) then %>		
		idBtnGuardar = tb.addButtonSAVE("Guardar", "canSubmit('<% =ACCION_GRABAR %>',0)");
		idBtnControl = tb.addButtonCONFIRM("Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");	
		<%	end if%>
		<% if fromTC <> 1 then %> 
			tb.addButton("Setting_folder-16x16.png", "Ajustes", "irAjustes()");	
		<% else %> 
			tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		<% end if %> 
		tb.draw();
		<%	
		index = 0	
		dim auxAju	
		if (not isFormSubmit()) then		
			if estaPMReferencia then
				while (readNextArticuloDB())
					Call PM2VS_DET()
					auxAju = getTotalArticuloxVale(idPMReferencia, vs_idarticulo, CODIGO_VS_AJUSTE_PEDIDO)
					%>
					myMS = agregarLineaArticulo();
					fillArticulo(myMS, <% =index %>, '<% =PM_idArticulo %>', '<% =PM_dsArticulo %>', <%=(cdbl(PM_Cantidad) - cdbl(auxAju))%>, <%=PM_Saldo%>, <% =(cdbl(PM_Cantidad) - cdbl(auxAju)) %>, '<% =PM_abreviaturaUnidad %>');					
					<% 
					index = index + 1
				wend			
			end if
		else
			if estaPMReferencia and idPMReferencia<>idPMReferenciaHDDN then
				PM_hayCabecera = True
				call initArticulosDB(idPMReferencia)
 				while (readNextArticuloDB())
					Call PM2VS_DET()
					auxAju = getTotalArticuloxVale(idPMReferencia, vs_idarticulo, CODIGO_VS_AJUSTE_PEDIDO)
					%>
					myMS = agregarLineaArticulo();
					fillArticulo(myMS, <% =index %>, '<% =PM_idArticulo %>', '<% =PM_dsArticulo %>', <%=(cdbl(PM_Cantidad) - cdbl(auxAju))%>, <%=PM_Saldo%>, <% =(cdbl(PM_Cantidad) - cdbl(auxAju)) %>, '<% =PM_abreviaturaUnidad %>');					
					<%
					index = index + 1
				wend			
			else
				PM_hayCabecera = True
				if initArticulosDB(idPMReferencia) then
					while (readNextArticuloVale(idVale))
 						%>
						myMS = agregarLineaArticulo();
						fillArticulo(myMS, <% =index %>, '<% =VS_idArticulo %>', '<% =VS_dsArticulo %>', <%=VS_Cantidad%>, <%=VS_Saldo%>, <% =VS_Cumplido %>, '<% =VS_abreviaturaUnidad %>');
						<%
						index = index + 1
					wend
				end if
			end if
		end if	
		%>
		pngfix();
		resaltarArticulosConErrores();	
	}
	
	function SeleccionarCalEmision(cal, date) {
		var str= new String(date);		
		document.getElementById("issuedateDiv").innerHTML = str;
	    document.getElementById("issuedate").value = str;
		if (cal) cal.hide();	
	}
	

	function CerrarCal(cal) {
		cal.hide();
	}
	
	function MostrarCalendario(p_objID, funcSel) {
		var dte= new Date();		    	    
		var elem= document.getElementById(p_objID);
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}	
	
	function resaltarArticulosConErrores(){
		//resalta con otro color articulos con errores, que consigue del array arrArticulosConErrores
		var tblArticulos = document.getElementById("tblArticulos");
		<%
		dim iArticulos
		For iArticulos = 0 to ubound(arrArticulosConErrores)%>
	      for (i=0; i< <%=index%>;i++){
			if (document.getElementById('item' + i ).value == '<%=arrArticulosConErrores(iArticulos)%>') tblArticulos.rows[i+1].className = 'reg_Header_Error';
		  }
	    <%next%>
	}
</script>
</head>

<script>
</script>
</head>
<body onLoad="bodyOnLoad()">	
<div id="toolbar"></div>
<br>		
<%
if (fromTC = 1) then
	submitPage = "almacenValesAJP.asp"
else
	submitPage = "almacenValesTitulo.asp?cdVAle=" & CODIGO_VS_AJUSTE_PEDIDO
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
			<td align="right" class="reg_Header_nav" colspan="5">&nbsp;</td>
		</tr>
		<tr>
			<td class="reg_Header_navdos"><%= GF_TRADUCIR("Nro Pedido Materiales") %></td>		
			<td colspan="2">
				<input type="text" id="pmReferencia" name="pmReferencia" value="<% =idPMReferencia %>" onChange="submitInfo('<%=ACCION_SUBMITIR%>');" size=5>
				<input type="hidden" id="pmReferenciaHDDN" name="pmReferenciaHDDN" value="<% =idPMReferencia %>">
			</td>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha Ajuste") %></td>
			<td width="15%"  align=center>
				<div id="issuedateDiv"><% =VS_FechaSolicitud %></div>
				<input type="hidden" id="issuedate" name="issuedate" value="<% =VS_FechaSolicitud %>">
			</td>
			<td colspan="1">				
				<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCalEmision)"><img id="imgEmision" src="images/DATE.gif"></a>			
			</td>			
		
		</tr>
		<input type="hidden" name="idObra" id="idObra" value="<% =vs_idObra %>">
		<tr>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Responsable") %></td>
			<td colspan="2">
				<%	response.write VS_dsSolicitante	%>															
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =VS_cdSolicitante %>"/>
			</td>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
			<td colspan="2">
				<%
				Set rsAlmacenes = obtenerListaAlmacenes(vs_idAlmacen) 
				if (not rsAlmacenes.eof) then 
					vs_idAlmacen = rsAlmacenes("IDALMACEN")
					response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
				end if
				%>
				<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =vs_idAlmacen %>">
			</td>
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
							<td align="center" width="7%"><% =GF_TRADUCIR("Codigo") %></td>
							<% if idVale = 0 then %>
								<td align="center" width="51%"><% =GF_TRADUCIR("Descripcion") %>	</td>
								<td align="center" width="14%"><% =GF_TRADUCIR("Cantidad Orig")%>	</td>
								<td align="center" width="14%"><% =GF_TRADUCIR("Entregado")%>		</td>
								<td align="center" width="14%"><% =GF_TRADUCIR("Nueva Cantidad") %>	</td>							
							<% else %>
								<td align="center" width="74%"><% =GF_TRADUCIR("Descripcion") %></td>						
								<td align="center" width="16%"><% =GF_TRADUCIR("Nueva Cantidad")%></td>
								<!--<td align="center" width="1%"></td>-->
							<% end if %>
						</tr>
					</table>
				</td>
			</tr>
		<% end if %>

	</table>
	<input type="hidden" id="accion" name="accion" value="">
	<input type="hidden" id="idVale" name="idVale" value="<% =idVale %>">
	<input type="hidden" id="cdVale" name="cdVale" value="<% =VS_cdVale %>">
	<input type="hidden" id="TC" name="TC" value="<% =fromTC %>">
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">
	<input type="hidden" name="resp" id="resp" value="MAYBE">		
</form>
</body>
</html>
<%
'---------------------------------------------------------------------------------------------
sub VS2PM()
	'PM = VS
	PM_FechaSolicitud = VS_FechaSolicitud
	PM_FechaRequerido = VS_FechaRequerido
	PM_cdSolicitante = VS_cdSolicitante
	PM_dsSolicitante = VS_dsSolicitante
	PM_idPedido = VS_idPedido
	PM_idAlmacen = VS_idAlmacen
	PM_idAlmacenDest = VS_idAlmacenDest	
	PM_idObra = VS_idObra
	'PM_idPedido = getIdPedidoCompleto()
	PM_idBudgetArea = VS_idBudgetArea
	PM_idBudgetDetalle = VS_idBudgetDetalle
	PM_usuario = VS_usuario
	PM_momento = VS_momento
	PM_hayCabecera = VS_hayCabecera
end sub
'---------------------------------------------------------------------------------------------
sub VS2PM_DET()
	PM_idArticulo = VS_idArticulo
	PM_dsArticulo = VS_dsArticulo
	PM_idUnidad = VS_idUnidad
	PM_abreviaturaUnidad = VS_abreviaturaUnidad
	PM_cantidad = VS_cantidad
	PM_saldo = VS_saldo
end sub
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