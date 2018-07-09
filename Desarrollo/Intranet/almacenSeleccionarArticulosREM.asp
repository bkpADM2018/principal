<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="includes/procedimientosObras.asp"-->
<!--#include file="includes/procedimientosPaginacion.asp"-->
<!--#include file="includes/procedimientosREM.asp"-->
<%

Const MAXLPP=50
'-----------------------------------------------------------------------------------------------
'Arma la lista con todos los PIC que tienen artículos no recibidos y que pertenecen a la division indicada
Function obtenerListaPICs(pIdDivison, pIdProveedor)
	
	Dim strSQL, rs, conn, strSQL1, strSQL2
	
	strSQL1=""
	strSQL2=""
	if (pIdProveedor <> 0) then
		strSQL1= " inner join TBLREMCABECERA RC on RC.IDREMITO=RP.IDREMITO where RC.IDPROVEEDOR=" & pIdProveedor
		strSQL2= " and CAB.IDPROVEEDOR=" & pIdProveedor
	end if
	
	'Se leen todo los PICs por cumplir, con todo o algo pendiente. Se excluyen los registros
	'con IDREMITO = 0, esto se hizo así como una marca para los PICs viejos que no deben aparecer.
	strSQL="			Select Distinct PEDIDO.IDCOTIZACION, CAB.IDPROVEEDOR, CAB.FECHAENTREGA, OBR.ESINVERSION"
	strSQL= strSQL & "	from ("
	strSQL= strSQL & "		Select PIC.* from TBLCTZCABECERA PIC"
	strSQL= strSQL & "		inner join 	(Select IDCOTIZACION IDPIC from TBLCTZCABECERA"
	strSQL= strSQL & "					EXCEPT"
	strSQL= strSQL & "					Select IDPIC from TBLREMPIC where IDREMITO=0) NPIC on PIC.IDCOTIZACION=NPIC.IDPIC"
	strSQL= strSQL & "		) CAB inner join  "
	strSQL= strSQL & "            (select sum(cantidad) cantidad,idcotizacion,idarticulo from TBLCTZDETALLE group by idcotizacion,idarticulo )"  
	strSQL= strSQL & "		PEDIDO on CAB.IDCOTIZACION=PEDIDO.IDCOTIZACION"
    strSQL= strSQL & "	left join  ("
						'Se obtiene todo lo recibido de un articulo para un PIC.
    strSQL= strSQL & "	    Select RP.IDPIC, RP.IDARTICULO, sum(RP.CANTIDAD) CANTIDAD"
	strSQL= strSQL & "	    from TBLREMPIC RP " & strSQL1
	strSQL= strSQL & "	    group by RP.IDPIC, RP.IDARTICULO"
    strSQL= strSQL & "	) RECIBIDO on PEDIDO.IDCOTIZACION = RECIBIDO.IDPIC and PEDIDO.IDARTICULO=RECIBIDO.IDARTICULO"
						'Se obtienen datos complemenarios para los articulos.
    strSQL= strSQL & "  INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO = PEDIDO.IDARTICULO"
						'Se obtienen datos complemenarios para las categorias.
	strSQL= strSQL & "  INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA = CAT.IDCATEGORIA"
						'Se obtienen datos complementarios de las obras.
	strSQL= strSQL & "  LEFT JOIN TBLDATOSOBRAS OBR on CAB.IDOBRA=OBR.IDOBRA"
	strSQL= strSQL & "	where (PEDIDO.CANTIDAD > RECIBIDO.CANTIDAD or RECIBIDO.CANTIDAD is null)"
	strSQL= strSQL & "		and (CAB.ESTADO='" & CTZ_FIRMADA & "' or CAB.ESTADO='" & CTZ_FACTURADA & "')"
	strSQL= strSQL & "		and CAT.TIPOCATEGORIA='" & TIPO_CAT_BIENES & "'"
	'strSQL= strSQL & "		and ART.BIENUSO <>'" & ES_BIEN_DE_USO & "'"
	strSQL= strSQL & "		and CAB.IDDIVISION=" & pIdDivison & " AND PEDIDO.cantidad > 0 " &strSQL2
	strSQL= strSQL & "	order by PEDIDO.IDCOTIZACION"
	'response.write strSQL
	'response.end
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaPICs = rs
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim aEntregar, rsSectores, conn, strSQL
Dim params, myIdAlmacen, idDivision
Dim rsComentarios, cdUsuario
dim rsAlmacenes, rs, lineasTotales, classPics

Set rsAlmacenes = obtenerListaAlmacenesUA()
if (rsAlmacenes.eof) then response.redirect "comprasAccesoDenegado.asp"
'Recibo los parametros.
pIdAlmacen = GF_PARAMETROS7("idAlmacen", 0, 6)
if (pIdAlmacen = 0) then pIdAlmacen = rsAlmacenes("IDALMACEN")
idProveedor = GF_PARAMETROS7("idProveedor",0,6)
dsProveedor = GF_PARAMETROS7("dsProveedor","",6)
verPorArticulo = GF_PARAMETROS7("vart", 0, 6)
if (idProveedor = 0) then verPorArticulo = 0
call addParam("vart", verPorArticulo, params)
call addParam("idProveedor", idProveedor, params)
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
call addParam("idAlmacen", idAlmacen, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10

'Determino la division de trabajo
idDivision= getDivisionAlmacen(pIdAlmacen)

'Obtengo la lista de articulos
if (verPorArticulo = 0) then
	Set aEntregar = obtenerListaPICs(idDivision, idProveedor)	
else
	Set aEntregar = obtenerArticulosPedidosNoRecibidos(idDivision, idProveedor, "", "", "")
end if
Call setupPaginacion(aEntregar, paginaActual, mostrar) 
lineasTotales = aEntregar.recordcount
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.Inversion {
	BACKGROUND-COLOR: #8CA4BC;
}

.Mantenimiento {
	BACKGROUND-COLOR: #F9D071;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<SCRIPT type="text/Javascript">
var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");

var FIRST_ROW = 2;

var logicalIndex = 0;
var nextRow = FIRST_ROW;	//Fila inicial de seleccion de articlos.

function seleccionarArticulo(indice) {	
	var idArticuloTemp = document.getElementById("idArticulo" + indice).value;
	var dsArticuloTemp = document.getElementById("dsArticulo" + indice).value;	
	var saldoTemp = document.getElementById("saldoArticulo" + indice).value.split(" ");
	
	var row = document.getElementById("tblSeleccion").insertRow(nextRow);
	row.id = "row" + logicalIndex;
	
	var colIdArticulo = row.insertCell(0);	
	colIdArticulo.innerHTML= idArticuloTemp;
	colIdArticulo.width="10%";
	colIdArticulo.align='center';
	
	var colDsArticulo = row.insertCell(1);
	colDsArticulo.innerHTML= dsArticuloTemp;
	
	
	var colSaldo = row.insertCell(2);
	colSaldo.align='right';
	colSaldo.innerHTML= saldoTemp[0];	
	
	var colUnit = row.insertCell(3);
	colUnit.width="30px";
	colUnit.innerHTML= saldoTemp[1];	
	
	var colDel = row.insertCell(4);
	colDel.align='center';
	var imgDelete = document.createElement("img");
	imgDelete.src = "images/compras/close-16x16.png";
	if (isFirefox) {
		imgDelete.setAttribute('onclick', "liberarArticulo(" + logicalIndex + ");");
	} else {
		imgDelete['onclick']=new Function("liberarArticulo(" + logicalIndex + ");return false;");
	}	
	imgDelete.style.cursor= "pointer";
	colDel.appendChild(imgDelete);
	
	nextRow++;
	logicalIndex++;
}

function liberarArticulo(indice) {	

	var row = document.getElementById("row" + indice)	
	if (row) {
		document.getElementById("tblSeleccion").deleteRow(row.rowIndex);
		nextRow--;
	}
}

function seleccionarProveedor(ms) {				
	var desc = ms.getSelectedItem();
	if (desc.indexOf('-') != -1) {
		var arr = desc.split('-');
		document.getElementById("idProveedor").value = arr[0];
		document.getElementById("dsProveedor").value = arr[1];
		ms.setValue(arr[1]);
	} else {
		if (desc == "") document.getElementById("idProveedor").value = 0;							
		if (desc == "") document.getElementById("dsProveedor").value = "";	
		ms.setValue("");
	}		
}
		
function lightOn(tr) {
	tr.className = "reg_Header_navdosHL";
}
	
function lightOff(tr, cls) {
	tr.className = cls;
}
	
function submitInfo(acc) {		
	document.getElementById("frmBusqueda").submit();
}
	
function irRemitos() {	
	location.href = "almacenAdministrarREM.asp";
}

function irCargaRemito(idPIC) {
	var flag = true;
	var param = "";
	if (idPIC > 0) {		
		param = "&ref=" + idPIC;
	} else {
		//Se carga por artículos		
		if (nextRow != FIRST_ROW) {
			var tb = document.getElementById("tblSeleccion");
			var index = 0
			for (var i = FIRST_ROW; i < nextRow; i++) {
				param += "&item" + index + "=" + tb.rows[i].cells[0].innerHTML
				param += "&amount" + index + "=" + tb.rows[i].cells[2].innerHTML;				
				index++;
			}		
			param += "&cantArticulos=" + index + "&idProveedor=<% =idProveedor %>&accion=<% =ACCION_SUBMITIR %>";
		} else {
			alert("<% =GF_TRADUCIR("Debe seleccinar al menos un articulo a recibir.") %>");
			flag = false;
		}
	}
	if (flag) location.href= "almacenREMTitulo.asp?cdREM=<% =CODIGO_REM_REMITO %>&idAlmacen=<% =pIdAlmacen %>" + param; 
}

function abrirCotizacion(idCTZ) {
	window.open("comprasPICPrint.asp?idCotizacionElegida=" + idCTZ, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);				
}
	
function startMagicSearch() {				
	var msProveedor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
	msProveedor.setMinChar(3);
	msProveedor.setToken(";");
	msProveedor.onBlur = seleccionarProveedor;		
	msProveedor.setValue('<%=dsProveedor%>');
}
	
function bodyOnLoad() {
	var tb = new Toolbar("toolBarGrupos",6, 'images/almacenes/');	
	tb.addButtonRETURN("Volver", "irRemitos()");
	tb.addButtonREFRESH("Recargar", "location.reload();");	
	tb.draw();	
	startMagicSearch();				
	<%	if (not aEntregar.eof) then		%>
			var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "almacenSeleccionarArticulosREM.asp<% =params %>");
	<%	end if %>
}
</script>

</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Nuevo Remito: Selección de Artículos") %>	
	<div id="toolBarGrupos"></div>
	<br>
	<form name="frmBusqueda" id="frmBusqueda" method="GET">
	<div id="busqueda">
	<table width="80%" cellspacing="0" cellpadding="0" align="center" border="0">
       <input type="hidden" name="accion" id="accion" value="">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Búsqueda") %></font></td>
           <td width="8"><img src="images/marco_r2_c3.gif"></td>
           <td></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="images/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="100%" align="center" border="0">
							<tr>
								<td align="right"><% = GF_TRADUCIR("Proveedor") %>:</td>
								<td>
									<div id="companyName0"></div>												
									<input type="hidden" id="idProveedor" name="idProveedor" value="<% =idProveedor %>">
									<input type="hidden" id="dsProveedor" name="dsProveedor" value="<% =dsProveedor %>">
								</td>
								<td align="right"><% =GF_TRADUCIR("Almacen") %>:</td>								
                                <td>                                
									<select id="idAlmacen" name="idAlmacen">
										<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -
										<%	
										while (not rsAlmacenes.eof)	%>
											<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = pIdAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
											<%		
											rsAlmacenes.MoveNext()
										wend 	
										%>		
									</select>		
                                </td>	
                            </tr>					
						<%	if (idProveedor <> 0) then %>
							<tr>
								<td align="right"><% =GF_TRADUCIR("Vista") %>:</td>	
								<td>
									<input type="radio" id="vart" name="vart" value="0" <% if (verPorArticulo = 0) then Response.Write "checked='checked'" %> >&nbsp;<% =GF_TRADUCIR("x PICs")%>
									<input type="radio" id="vart" name="vart" value="1" <% if (verPorArticulo = 1) then Response.Write "checked='checked'" %> >&nbsp;<% =GF_TRADUCIR("x Articulos") %>
								</td>
							</tr>	
						<%	end if			%>
                            <tr>
								<td colspan="4" align="center"><input type="button" value="Buscar..." id="submit1" name="submit1" onclick='submitInfo();'></td>
                            </tr>								
                     </table>
	           </td>
	           <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="images/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="images/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	</div>	
	</form>
	<br>
	<form name="frmSel" id="frmSel" method="POST" method="almacenREMTitulo.asp">
	<%	if (verPorArticulo = 1) then %>
		<table id="tblSeleccion" class="reg_header" cellpadding="2" cellspacing="1" width="80%" align="center">
			<tr class="reg_header_nav"><td align="center" colspan="5"><% =GF_TRADUCIR("Articulos Seleccionados") %></td></tr>
			<tr class="reg_header_nav">
				<td align="center" width="70%" colspan="2"><% =GF_TRADUCIR("Articulo") %></td>
				<td align="center" colspan="2"><% =GF_TRADUCIR("Saldo") %></td>
				<td align="center" width="5%">.</td>
			</tr>
			<tr class="reg_header_nav">
				<td align="center" colspan="5"><input type="button" value="<% =GF_TRADUCIR("Cargar Remito") %>" onClick="irCargaRemito(0)"></td>
			</tr>
		</table>
	<%	end if %>
	</form>
	<br>	
	<table align="center" width="80%" class="reg_Header">
			<% 	if (not aEntregar.eof) then %>
				<%	if (verPorArticulo = 1) then %>
					<tr><td colspan="3"><div id="paginacion"></div></td></tr>
					<tr class="reg_Header_nav">
						<td align="center"><% =GF_TRADUCIR("Articulo") %></td>
						<td align="center"><% =GF_TRADUCIR("Cantidad") %><br><% =GF_TRADUCIR("requerida") %></td>
						<td align="center"><% =GF_TRADUCIR("Cantidad") %><br><% =GF_TRADUCIR("recibida") %></td>
					</tr>
				<%	else	%>
					<tr><td colspan="4"><div id="paginacion"></div></td></tr>
					<tr>
						<td colspan="2">
							<table><tr>
								<td class="Inversion">&nbsp;&nbsp;&nbsp;&nbsp;</td><td><% =GF_TRADUCIR("OBRAS DE INVERSION") %></td>
								<td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td class="Mantenimiento">&nbsp;&nbsp;&nbsp;&nbsp;</td><td><% =GF_TRADUCIR("OBRAS DE MANTENIMIENTO") %></td>
								<td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td class="reg_Header_navdos">&nbsp;&nbsp;&nbsp;&nbsp;</td><td><% =GF_TRADUCIR("OTROS") %></td>
							</tr></table>
						</td>
					</tr>
					<tr class="reg_Header_nav">
						<td align="center"><% =GF_TRADUCIR("Nro. PIC") %></td>
						<td align="center"><% =GF_TRADUCIR("Proveedor") %></td>
						<td align="center"><% =GF_TRADUCIR("Fecha") %><br><% =GF_TRADUCIR("Entrega") %></td>
						<td align="center">.</td>
					</tr>
				<%	end if %>
			<%	end if %>
<%	reg=0
	if (not aEntregar.eof) then	
		intIndice=0	
		while ((not aEntregar.eof) and (CInt(intIndice) < CInt(mostrar)))			
			intIndice= intIndice+1
			classPics = ""
			if (verPorArticulo = 0) then
				if (aEntregar("ESINVERSION") = OBRA_INVERSION) then
					classPics = " Inversion"
				elseif (aEntregar("ESINVERSION") = OBRA_MANTENIMIENTO) then
					classPics = " Mantenimiento"
				end if
			end if
%>			
			<tr class="reg_Header_navdos<% =classPics %>" style="cursor:pointer" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this, 'reg_Header_navdos<% =classPics %>')">
				<%	if (verPorArticulo = 1) then %>							
				<td align="left" onClick="javascript:seleccionarArticulo(<% =intIndice %>)"><% =aEntregar("idarticulo")%>-<% =aEntregar("DSARTICULO") %></td>				
				<td align="right" onClick="javascript:seleccionarArticulo(<% =intIndice %>)"><% =aEntregar("CANTIDADP") %>&nbsp;<% =aEntregar("UNIDAD") %></td>
				<td align="right" onClick="javascript:seleccionarArticulo(<% =intIndice %>)"><% if (aEntregar("CANTIDADR") <> "") then Response.write aEntregar("CANTIDADR") & "&nbsp;" & aEntregar("UNIDAD") %></td>				
				
				<input type="hidden" id="idArticulo<%=(intIndice)%>" value="<% =aEntregar("idarticulo") %>">
				<input type="hidden" id="dsArticulo<%=(intIndice)%>" value="<% =aEntregar("DSARTICULO") %>">
				<input type="hidden" id="saldoArticulo<%=(intIndice)%>" value="<% 
					if (aEntregar("CANTIDADR") <> "") then 
						Response.write CDbl(aEntregar("CANTIDADP"))-CDbl(aEntregar("CANTIDADR")) & " " & aEntregar("UNIDAD")
					else
						Response.Write aEntregar("CANTIDADP") & " " & aEntregar("UNIDAD")
					end if
				%>">
				<input type="hidden" id="unidad<%=(intIndice)%>" value="<% =aEntregar("UNIDAD") %>">							
				<%	else	%>			
				<td align="center" onClick="javascript:irCargaRemito(<% =aEntregar("IDCOTIZACION")%>)"><% =aEntregar("IDCOTIZACION")%></td>
				<td align="left" onClick="javascript:irCargaRemito(<% =aEntregar("IDCOTIZACION")%>)"><% =aEntregar("IDPROVEEDOR")%>-<% =getDescripcionProveedor(aEntregar("IDPROVEEDOR")) %></td>
				<td align="center" onClick="javascript:irCargaRemito(<% =aEntregar("IDCOTIZACION")%>)"><% =GF_FN2DTE(aEntregar("FECHAENTREGA")) %></td>
				<td align="center"><img src="images/almacenes/printer-16x16.png" onclick="javascript:abrirCotizacion(<% =aEntregar("IDCOTIZACION") %>)"></td>			
				<%	end if	%>			
			</tr>
	<%		aEntregar.MoveNext()					
		wend
	else	%>
			<tr class="TDNOHAY"><td colSpan="4"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
<%  end if %>				
		</table>		
</body>
</html>
<%
'******************************************************************************************
	Function addParam(p_strKey,p_strValue,ByRef p_strParam)
           if (not isEmpty(p_strValue)) then
              if (isEmpty(p_strParam)) then
                 p_strParam = "?"
              else
                 p_strParam = p_strParam & "&"
              end if
              p_strParam = p_strParam & p_strKey & "=" & p_strValue
           end if
	End Function
		
%>