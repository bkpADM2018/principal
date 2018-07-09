<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->

<%
'Call controlAccesoCM("CMADM")

'-----------------------------------------------------------------------------------------------
Function filtrarArticulos(ByRef myWhere, idcategoria, idarticulo, cdinterno, dsarticulo)
		
	'Filtro
	myWhere = ""
	if (idarticulo <> "") then  Call mkWhere(myWhere, "A.IDARTICULO", idarticulo, "=", 1)
	if (CLng(idcategoria) <> 0) then  Call mkWhere(myWhere, "A.IDCATEGORIA", idcategoria, "=", 1)
	if (cdinterno <> "") then  Call mkWhere(myWhere, "A.CDINTERNO", cdinterno, "LIKE", 3)	
	if (dsarticulo <> "") then  Call mkWhere(myWhere, "A.DSARTICULO", dsarticulo, "LIKE", 3)	
	filtrarArticulos = myWhere	
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerListaArticulos(idAlmacen, idcategoria, idarticulo, cdinterno, dsarticulo, pagina, regXpag) 
	Dim strSQL, rs, myWhere, firstRecord, conn
	
	Call filtrarArticulos(myWhere, idcategoria, idarticulo, cdinterno, dsarticulo)	
	
	strSQL= "Select T.IDARTICULO, T.DSARTICULO, T.IDDIVISION, T.DSDIVISION, T.DSCATEGORIA, case when STOCK is Null then 0 else STOCK end STOCK, abreviatura " 
	if (idAlmacen <> 0) then
		strSQL= strSQL & ", cdinterno, stockmaximo, stockminimo, compramaxima, compraminima"
	end if
	strSQL= strSQL & "	from (Select A.IDARTICULO, A.DSARTICULO, CAT.DSCATEGORIA, A.IDUNIDAD, D.IDDIVISION, DSDIVISION, sum(EXISTENCIA+SOBRANTE) STOCK " &_
					"			from TBLARTICULOS A inner join TBLDIVISIONES D on D.CDDIVISIONABR <> '" & CODIGO_EXPORTACION & "' and A.ESTADO <> " & ESTADO_BAJA &_
					"			inner join TBLARTCATEGORIAS CAT on A.IDCATEGORIA=CAT.IDCATEGORIA and TIPOCATEGORIA = '" & TIPO_CAT_BIENES & "'" &_
					"			left join TBLALMACENES AL on AL.IDDIVISION=D.IDDIVISION " &_
					"			left join TBLARTICULOSDATOS AD on AD.IDALMACEN=AL.IDALMACEN and AD.IDARTICULO=A.IDARTICULO " & myWhere &_
					"			group by A.IDARTICULO, A.DSARTICULO, CAT.DSCATEGORIA, A.IDUNIDAD, D.IDDIVISION, DSDIVISION " &_
					"		) T " &_
					"		left join TBLUNIDADES unidades on T.idunidad=unidades.idunidad "
	if (idAlmacen <> 0) then
		strSQL= strSQL & "   left join tblarticulosdatos as articulosdatos " &_
						" 	on T.idarticulo = articulosdatos.idarticulo and articulosdatos.idalmacen = " & idAlmacen
	end if
	strSQL= strSQL & "order by T.IDARTICULO, T.IDDIVISION "
	'response.write strSQL	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	Call setupPaginacion(rs, pagina, regXPag)
	
	Set obtenerListaArticulos = rs
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim articulos, idArticulo
Dim rsAlmacen, idCategoria
Dim  cdinterno, dsarticulo, stockmax, stockmin, compramax, compramin, stockunificado
Dim params
Dim reg, lineasTotales, paginaActual


idPedido = GF_PARAMETROS7("idPedido","",6)
call addParam("idPedido", idPedido, params)
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
call addParam("idAlmacen", idAlmacen, params)
idCategoria = GF_PARAMETROS7("idCategoria",0,6)
call addParam("idCategoria", idCategoria, params)
idarticulo = GF_PARAMETROS7("idarticulo","",6)
call addParam("idarticulo", idarticulo, params)
cdinterno = GF_PARAMETROS7("cdinterno","",6)
call addParam("cdinterno", cdinterno, params)
dsarticulo = UCase(GF_PARAMETROS7("dsarticulo","",6))
call addParam("dsarticulo", dsarticulo, params)
stockmax = GF_PARAMETROS7("stockmax",0,6)
call addParam("stockmax", stockmax, params)
stockmin = GF_PARAMETROS7("stockmin",0,6)
call addParam("stockmin", stockmin, params)
compramax = GF_PARAMETROS7("compramax",0,6)
call addParam("compramax", compramax, params)
compramin = GF_PARAMETROS7("compramin",0,6)
call addParam("compramin", compramin, params)
stockunificado = GF_PARAMETROS7("stockunificado",0,6)
call addParam("stockunificado", stockunificado, params)

paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 50
cdUsuario = ""


GP_ConfigurarMomentos

Set articulos = obtenerListaArticulos(idAlmacen, idCategoria, idarticulo, cdinterno, dsarticulo, paginaActual, mostrar)
lineasTotales = articulos.RecordCount


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=11">
<title>Sistema de Almacenes - Administrar Articulos</title>

<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
	function abrirArticulo(idAlmacen,idArticulo) {	
		if (idAlmacen != 0) {
			var puw = new winPopUp('popupArticulo',"comprasPropArticulo.asp?idAlmacen=" + idAlmacen + "&idArticulo=" + idArticulo,'500','450','Propiedades Articulo', 'submitInfo()');	
		}
	}
	
	function submitInfo(acc) {		
		document.getElementById("frmSel").submit();
	}
	
	function volver() {
		location.href = "almacenIndex.asp";
	}
	
	function irHome() {
		location.href = "almacenIndex.asp";
	}
		
	function irAdministracion() {
		location.href = "almacenAdministracion.asp";
	}
	
	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}
	
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, 'images/almacenes/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");		
		tb.addButtonREFRESH("Recargar", "submitInfo()");				
		tb.draw();		
		<%	
			if (not articulos.eof) then		%>								
				var pgn = new Paginacion("paginacion");							
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 200, "almacenAdministrarArticulosAlmacen.asp<% =params %>");
		<%	end if %>
	}
</script>
</head>
<body onLoad="bodyOnLoad()">	
	<div id="toolbar"></div>	
	<br>
	<form name="frmSel" id="frmSel">
	<div class="tableaside size100"> 
		<h3> Filtros </h3>
		  
		<div id="searchfilter" class="tableasidecontent">
	        
	        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("C&oacutedigo") %>: </div>
	        <div class="col16"> <input type="text" id="idArticulo" name="idArticulo" value="<% =idArticulo %>"></div>
	        	        		        
	        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("C&oacutedigo Interno") %>: </div>
	        <div class="col16"> <input type="text" id="cdInterno" name="cdInterno" value="<% =cdInterno %>"> </div>
		       
	        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Descripci&oacuten") %>: </div>
	        <div class="col16"> <input type="text" id="dsarticulo" name="dsarticulo" value="<% =dsarticulo %>"> </div>
			
			<% 
				Set rsAlmacen = obtenerListaAlmacenesUA()					
				if (not rsAlmacen.eof) then
			%>
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Almacen") %>: </div>
	        <div class="col16"> 				
				<select id="idAlmacen" name="idAlmacen">
					<option value="0" <% if (rsAlmacen("IDALMACEN") = idAlmacen) then response.write "selected='true'" %>> - <% =GF_TRADUCIR("Todas") %> - </option>
				<%	while (not rsAlmacen.eof) %>
						<option value="<% =rsAlmacen("IDALMACEN") %>" <% if (rsAlmacen("IDALMACEN") = idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacen("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacen("DSALMACEN")) %></option>
				<%		rsAlmacen.MoveNext()
					wend	%>
				</select>
			</div>
			<%	end if %>
			<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Categor&iacutea") %>: </div>
	        <div class="col16"> 
				<% 
					Call executeQueryDb(DBSITE_SQL_INTRA, rsCat, "OPEN", "Select IDCATEGORIA, DSCATEGORIA from TBLARTCATEGORIAS where TIPOCATEGORIA = '" & TIPO_CAT_BIENES & "' order by DSCATEGORIA")					
				%>
				<select id="idCategoria" name="idCategoria">
					<option value="0" <% if (rsCat("IDCATEGORIA") = idCategoria) then response.write "selected='true'" %>> - <% =GF_TRADUCIR("Todas") %> - </option>
				<%	while (not rsCat.eof) %>
						<option value="<% =rsCat("IDCATEGORIA") %>" <% if (rsCat("IDCATEGORIA") = idCategoria) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsCat("DSCATEGORIA")) %></option>
				<%		rsCat.MoveNext()
					wend	%>
				</select>
			</div>
			
	    	<span class="btnaction"><input type="submit" value="Buscar"></span>
		</div>
	</div>
    
	<div class="col66"></div>	
	</form>
	<br>
	<table align="center" class="datagrid" width="95%">
		<thead>			
			<tr>
				<th style="text-align: center"><% =GF_TRADUCIR("C&oacutedigo") %></th>								
				<th style="text-align: center"><% =GF_TRADUCIR("Descripci&oacuten") %></th>		
				<th style="text-align: center"><% =GF_TRADUCIR("Categor&iacutea") %></th>		
				<%	if (idAlmacen <> 0) then	%>
					<th style="text-align: center"><% =GF_TRADUCIR("C&oacutedigo<br>Interno") %></th>
					<th style="text-align: center"><% =GF_TRADUCIR("Stock<br>Maximo") %></th>
					<th style="text-align: center"><% =GF_TRADUCIR("Stock<br>Minimo") %></th>
					<th style="text-align: center"><% =GF_TRADUCIR("Compra<br>Maxima") %></th>
					<th style="text-align: center"><% =GF_TRADUCIR("Compra<br>Maxima") %></th>
				<%	end if	
					Call executeQueryDb(DBSITE_SQL_INTRA, rsDivi, "OPEN", "Select DSDIVISION from TBLDIVISIONES where CDDIVISIONABR<> '" & CODIGO_EXPORTACION & "' order by IDDIVISION")
					while (not rsDivi.eof)
				%>				
					<th style="text-align: center" width="7%"><% =GF_TRADUCIR("Stock<br>") & rsDivi("DSDIVISION") %></th>				
				<%		rsDivi.MoveNext()
					wend	%>
			</tr>	
		</thead>		
<%	reg=0
	if (not articulos.eof) then			%>
		<tbody>
<%			while ((not articulos.eof) and (reg < mostrar))	
				myIdArticulo = articulos("IDARTICULO")
				reg=reg+1
%>
			<tr style="cursor:pointer" onclick="abrirArticulo(<%=idAlmacen%>, <% =articulos("idArticulo") %>)">			
				<td style="text-align: center"><% =articulos("idarticulo") %>		</td>								
				<td style="text-align: left">  <% =articulos("dsarticulo") %>		</td>	
				<td style="text-align: center">  <% =articulos("dscategoria") %>		</td>					
				<%	if (idAlmacen <> 0) then	%>
					<td style="text-align: center"><% =articulos("cdinterno") %>		</td>	
					<td style="text-align: right"> <% =articulos("stockmaximo") %>&nbsp;<% if (articulos("stockmaximo") <> "") then Response.write articulos("abreviatura") %>		</td>				
					<td style="text-align: right"> <% =articulos("stockminimo") %>&nbsp;<% if (articulos("stockminimo") <> "") then Response.write articulos("abreviatura") %></td>				
					<td style="text-align: right"> <% =articulos("compramaxima") %>&nbsp;<% if (articulos("compramaxima") <> "") then Response.write articulos("abreviatura") %></td>				
					<td style="text-align: right"> <% =articulos("compraminima") %>&nbsp;<% if (articulos("compraminima") <> "") then Response.write articulos("abreviatura") %></td>				
				<%	end if	
					flagSalir = false
					while ((not articulos.eof) and (not flagSalir))	
						if (myIdArticulo = articulos("IDARTICULO")) then
					%>
							<td style="text-align: right"> <% =articulos("stock") & " " & articulos("abreviatura") %> </td>				
				<%			articulos.MoveNext()
						else
							flagSalir = true
						end if
					wend	%>
			</tr>		
	<%		wend 	%>
		</tbody>		
		<tfoot>
			<tr><td colspan="12"><div id="paginacion"></div></td></tr>
		</tfoot>
<%	end if
	if (reg = 0) then
%>
		<tbody>	<tr class="TDNOHAY" ><td colSpan="12"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		</tbody>
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