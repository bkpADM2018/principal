<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<%
Const COLUMNAS_CON_VLU = 15
Const COLUMNAS_SIN_VLU = 7

Const FILTRO_TODOS		= 0
Const FILTRO_EXISTENTES = 1

'-----------------------------------------------------------------------------------------------
Function getSQLStock()
	dim mywhere, strSQL
	strSQL = ""
	
	if (metodo = FILTRO_EXISTENTES) then mywhere = " WHERE (cie.existencia <> 0 OR cie.sobrante <> 0)"
	
	Call mkWhere(mywhere, "cie.idalmacen", almacen,"=", 1)

	strSQL = "SELECT			cie.idarticulo                          , "
	strSQL = strSQL & "         art.dsarticulo				            , "
	strSQL = strSQL & "         cat.idcategoria                         , "
	strSQL = strSQL & "         cat.dscategoria                         , "
	strSQL = strSQL & "         cie.existencia					        , "
	strSQL = strSQL & "         uni.abreviatura unidad                  , "
	strSQL = strSQL & "         cie.sobrante							, "
	strSQL = strSQL & "         cie.existencia + cie.sobrante total		, "
	strSQL = strSQL & "         cie.vlupesos						    , "
	strSQL = strSQL & "         cie.vludolares							, "
	strSQL = strSQL & "         cie.mmtocompra						    , "
	strSQL = strSQL & "         cie.vlupesoscompra					    , "
	strSQL = strSQL & "         cie.vludolarescompra					, "
	strSQL = strSQL & "         cie.idpic								, "
	strSQL = strSQL & "         ard.cdinterno                           , "
	strSQL = strSQL & "         art.MMTOULTIMACOMPRA mmtoUC             , "
	strSQL = strSQL & "         art.VLUPESOSULTIMACOMPRA valorPesosUC   , "
    strSQL = strSQL & "         art.VLUDOLARESULTIMACOMPRA valorDolaresUC , "
    strSQL = strSQL & "         art.idpic picUC "
	strSQL = strSQL & "FROM     TBLREPORTESTOCKWF cie "
	strSQL = strSQL & "         INNER JOIN TBLARTICULOS art "
	strSQL = strSQL & "				ON	cie.idarticulo  = art.idarticulo "
	strSQL = strSQL & "				AND cie.cdusuario = '" & session("Usuario") & "'"
	strSQL = strSQL & "			INNER JOIN TBLARTCATEGORIAS cat "
	strSQL = strSQL & "				ON  art.idcategoria = cat.idcategoria "
	strSQL = strSQL & "         INNER JOIN TBLUNIDADES uni "
	strSQL = strSQL & "				ON  art.idunidad = uni.idunidad "
	strSQL = strSQL & "         LEFT JOIN TBLARTICULOSDATOS ard "
	strSQL = strSQL & "				ON  ard.idArticulo = cie.idArticulo "
	strSQL = strSQL & "				AND ard.idAlmacen = " & almacen
	strSQL = strSQL & mywhere
	strSQL = strSQL & "         ORDER BY art.idarticulo"	
	call logdebug(strSQL)
	getSQLStock = strSQL
	
End Function
'-----------------------------------------------------------------------------------------------
Function getArticuloVLU()	
		
	XLS_flagUCTablaArt = false
	
	XLS_idCategoria = rsStock("IDCATEGORIA")
	XLS_dsCategoria = trim(rsStock("DSCATEGORIA"))
	XLS_idArticulo = rsStock("IDARTICULO")
	XLS_dsArticulo = trim(rsStock("DSARTICULO"))
	XLS_unidad = trim(rsStock("UNIDAD"))
	XLS_ubicacion = trim(rsStock("CDINTERNO"))
	
	if (isNull(rsStock("EXISTENCIA")))then XLS_existencia = 0 else XLS_existencia = cdbl(rsStock("EXISTENCIA"))  end if
	if (isNull(rsStock("SOBRANTE")))  then XLS_sobrante   = 0 else XLS_sobrante = cdbl(rsStock("SOBRANTE"))  end if
	if (isNull(rsStock("TOTAL")))     then XLS_total      = 0 else XLS_total = cdbl(rsStock("TOTAL"))  end if
	if (valorizar) then 

		if (not isNull(rsStock("VLUPESOS"))) then
			XLS_vluPesos = cdbl(rsStock("VLUPESOS"))
			XLS_vluPesosTotal = GF_EDIT_DECIMALS(XLS_existencia * XLS_vluPesos, 2)				
			XLS_vluPesos = GF_EDIT_DECIMALS(XLS_vluPesos, 2)
		else
			XLS_vluPesos = GF_EDIT_DECIMALS("000", 2)
			XLS_vluPesosTotal = XLS_vluPesos
		end if
		
		if (not isNull(rsStock("VLUDOLARES"))) then
			XLS_vluDolar = CDbl(rsStock("VLUDOLARES"))
			XLS_vluDolarTotal = GF_EDIT_DECIMALS(XLS_existencia * XLS_vluDolar, 2)				
			XLS_vluDolar = GF_EDIT_DECIMALS(XLS_vluDolar, 2)
		else
			XLS_vluDolar = GF_EDIT_DECIMALS("000", 2)
			XLS_vluDolarTotal = XLS_vluDolar
		end if		

		'Ultima compra -----------------------------
		
		if (isNull(rsStock("IDPIC")))		 then 
			if (CLng(rsStock("PICUC")) <> 0) then
				XLS_flagUCTablaArt = true
				'Se asume que si hay uno de los datos de ultima compra => estan todos, sino hay un error.
				XLS_fechaUltimaCompra = GF_FN2DTE(left(rsStock("mmtoUC"),8))
				XLS_vluPesosUltimaCompra = GF_EDIT_DECIMALS(CLng(rsStock("valorPesosUC")),2)
				XLS_vluDolaresUltimaCompra = GF_EDIT_DECIMALS(CLng(rsStock("valorDolaresUC")),2)	
				XLS_idPIC = rsStock("PICUC")
			else
				XLS_fechaUltimaCompra = ""		
				XLS_vluPesosUltimaCompra = ""		
				XLS_vluDolaresUltimaCompra = ""	
				XLS_idPIC = ""
			end if
		else	
			'Se asume que si hay uno de los datos de ultima compra => estan todos, sino hay un error.
			XLS_fechaUltimaCompra = GF_FN2DTE(left(rsStock("MMTOCOMPRA"),8))
			XLS_vluPesosUltimaCompra = GF_EDIT_DECIMALS(CLng(rsStock("VLUPESOSCOMPRA")),2)
			XLS_vluDolaresUltimaCompra = GF_EDIT_DECIMALS(CLng(rsStock("VLUDOLARESCOMPRA")),2)
			XLS_idPIC = rsStock("IDPIC")
		end if
		
	end if
	
	
End Function
'-----------------------------------------------------------------------------------------------
Function resetArticuloVlu()
	XLS_idCategoria = ""
	XLS_dsCategoria = ""
	XLS_idArticulo = ""
	XLS_dsArticulo = ""
	XLS_unidad = ""
	XLS_ubicacion = ""
	XLS_existencia = ""
	XLS_sobrante = ""
	XLS_total = ""
	XLS_vluPesos = ""
	XLS_vluDolar = ""
	XLS_vluPesosTotal = ""
	XLS_vluDolarTotal = ""
	XLS_fechaUltimaCompra = ""
	XLS_vluPesosUltimaCompra = ""
	XLS_vluDolaresUltimaCompra = ""
	XLS_idPIC = ""
End Function
'-----------------------------------------------------------------------------------------------
Function dibujarEncabezado()
	dim strSQL, rs1, conn, rs2, auxcol
	auxcol = COLUMNAS_SIN_VLU
	if (valorizar) then	auxcol = COLUMNAS_CON_VLU
	%>
	<table class="xls_border_left">
		<tr><td colspan="<% =auxcol %>"></td><td colspan="2" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =auxcol %>" align="center" style="font-size:24">STOCK DE ARTICULOS</td></tr>
	</table>
	<%
	if (categoria <> -1) then
		strSQL = "select idcategoria id,dscategoria ds from tblartcategorias where idcategoria = " & categoria
		Call executeQueryDB(DBSITE_SQL_INTRA, rs1, "OPEN", strSQL)
		if (not rs1.eof) then dscategoria= rs1("ds")
	else
		dscategoria="Todas"
	end if
	strSQL = "select idalmacen id,dsalmacen ds from tblalmacenes where idalmacen = " & almacen		
	Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
	if (not rs2.eof) then dsalmacen=rs2("ds")
	if (metodo = FILTRO_TODOS) then
		busqueda="Todos"
	else
		busqueda="Con Stock"
	end if
	fecha = trim(GF_FN2DTE(fechaBusqueda))
	%>
	<table style="font-size:16; font-weight:bold; font-family:courier">
	<tr><td></td></tr>
	<tr><td>Categoria.:		</td><td><% =dscategoria	%></td></tr>
	<tr><td>Almacen...:	</td><td><% =dsalmacen		%></td></tr>
	<tr><td>Buscar....:	</td><td><% =busqueda		%></td></tr>
	<tr align="left"><td>Stocks al.:	</td><td><% =fecha %></td><td colspan="6">
	<% 
	if incluir then 
		response.write "(Datos calculados al final de la fecha seleccionada.)"
	else
		response.write "(Datos calculados al inicio de la fecha seleccionada.)"
	end if
	%>
	</td></tr>
	<tr><td>Valorizar.: </td><td> <% if (valorizar) then Response.Write "Si" else Response.Write "No" end if %></td></tr>
	<tr><td></td></tr>
	<tr><td><u>Leyendas</u></td></tr>
	<tr><td></td></tr>
	<tr><td class="xls_precioUC_tablaArticulos"></td><td colspan=8 style="font-size:12px">El importe no pertenece a la división consultada.</td></tr>
	<tr><td></td></tr>
	</table><%
End Function
'-----------------------------------------------------------------------------------------------

'******************************************************
'					INICIO DE LA PAGINA
'******************************************************
dim accion, metodo, categoria, almacen, valorizar, fechaBusqueda, dscategoria
dim conn, rsStock, strSQL, fname, conFecha, KKK, fecha, busqueda, dsalmacen
dim XLS_idCategoria, XLS_dsCategoria, XLS_idArticulo, XLS_dsArticulo, XLS_unidad
dim XLS_existencia, XLS_sobrante, XLS_total, XLS_vluPesos, XLS_vluDolar, XLS_ubicacion
dim XLS_vluPesosTotal, XLS_vluDolarTotal, XLS_fechaUltimaCompra, XLS_vluPesosUltimaCompra, XLS_vluDolaresUltimaCompra,XLS_idPIC
dim XLS_flagImporteGeneral, XLS_CSSWarning, XLS_flagUCTablaArt,CssUCTablaArt

accion    = GF_PARAMETROS7("accion"   ,"", 6)
metodo    = GF_PARAMETROS7("metodo"   ,0 , 6)
categoria = GF_PARAMETROS7("categoria",0 , 6)
almacen   = GF_PARAMETROS7("almacen"  ,0 , 6)
valorizar = GF_PARAMETROS7("valorizar","", 6)
fechaBusqueda = GF_PARAMETROS7("fechaBusqueda", "", 6)
fechaBusqueda = GF_DTE2FN(fechaBusqueda)
incluir = GF_PARAMETROS7("incluir", "", 6)
if (valorizar = "on") then valorizar = true

if (accion = ACCION_PROCESAR) then
	fname = session("usuario") & "_" & almacen & "_" & fechaBusqueda
	strSQL = getSQLStock()
	call executeQueryDb(DBSITE_SQL_INTRA, rsStock, "OPEN", strSQL)
	Call GF_createXLS(fname)
	%>
	<html>
	<head>
		<style type="text/css">
			.xls_border_left { 
				border-color:#666666; 
				border-style:solid; 
				border-width:thin;
			}
			.xls_align_center { 
				border-color:#666666; 
				border-style:solid; 
				border-width:thin;
				text-align: center;
			}
			.xls_align_right { 
				border-color:#666666; 
				border-style:solid; 
				border-width:thin;
				text-align: right;
			}
			.xls_precioUC_tablaArticulos
			{
				BACKGROUND-COLOR: #ffff80;
				border-color:#666666; 
				border-style:solid; 
				border-width:thin;
			}
		</style>
	</head>
	<body>
	<table class="xls_border_left" style="background-color:#FFFACD; font-weight:bold">
		<tr><td><% Call dibujarEncabezado() %></td></tr>
	</table>
	<table class="xls_border_left" style="background-color:#E3F6CE; font-weight:bold">
		<tr>
			<td colspan="2" class="xls_align_center">CATEGORIAS</td>
			<td colspan="4" class="xls_align_center">ARTICULOS</td>
			<td colspan="3" class="xls_align_center">STOCK</td>
			<% if (valorizar) then  %>
					<td colspan="4" class="xls_align_center">VALUACION</td>
					<td colspan="4" class="xls_align_center">ULTIMO COMPRA</td>
			<% end if %>
		</tr>
	</table>
	<table  class="xls_border_left" style="background-color:#E0E0F8; font-weight:bold">
		<tr>
			<td class="xls_align_center">ID CATEGORIA</td>
			<td class="xls_border_left">DESCRIPCIÓN</td>
			<td class="xls_align_center">ID ARTICULO</td>
			<td class="xls_align_center">DESCRIPCIÓN</td>
			<td class="xls_align_center">UNIDAD</td>
			<td class="xls_align_center">UBICACION</td>
			<td class="xls_align_center">EXISTENCIA</td>
			<td class="xls_align_center">SOBRANTE</td>
			<td class="xls_align_center">TOTAL</td>
			<% if (valorizar) then  %>
					<td class="xls_align_center">VALOR $/U</td>
					<td class="xls_align_center">VALOR U$S/U</td>
					<td class="xls_align_center">VALUACION $</td>
					<td class="xls_align_center">VALUACION U$S</td>
					<td class="xls_align_center">FECHA</td>
					<td class="xls_align_center">PRECIO $/U</td>
					<td class="xls_align_center">PRECIO U$S/U</td>
					<td class="xls_align_center">PIC</td>
			<% end if %>
		</tr>
	</table>
	<table class="xls_border_left">
		<%
		While (not rsStock.eof) %>
				<% Call getArticuloVlu() 
				
				CssUCTablaArt = ""
				if (XLS_flagUCTablaArt) then CssUCTablaArt = "xls_precioUC_tablaArticulos"
				%>
				<tr>
					<td class="xls_align_center">	<% =XLS_idCategoria %></td>
					<td class="xls_border_left">	<% =XLS_dsCategoria %></td>
					<td class="xls_align_center">	<% =XLS_idArticulo %></td>
					<td class="xls_border_left">	<% =XLS_dsArticulo %></td>
					<td class="xls_align_center">	<% =XLS_unidad %></td>
					<td class="xls_align_center">	<% =XLS_ubicacion %></td>
					<td class="xls_align_right">	<% =XLS_existencia %></td>
					<td class="xls_align_right">	<% =XLS_sobrante %></td>
					<td class="xls_align_right">	<% =XLS_total %></td>										
					<% if (valorizar) then  %>
							<td class="xls_align_right">	<% =XLS_vluPesos %></td>
							<td class="xls_align_right">	<% =XLS_vluDolar %></td>
							<td class="xls_align_right">	<% =XLS_vluPesosTotal %></td>
							<td class="xls_align_right">	<% =XLS_vluDolarTotal %></td>
							<td class="<%=CssUCTablaArt%>" align="center" style="border:thin solid #666666">	<% =XLS_fechaUltimaCompra %></td>
							<td class="<%=CssUCTablaArt%>" align="right" style="border:thin solid #666666">	<% =XLS_vluPesosUltimaCompra %></td>
							<td class="<%=CssUCTablaArt%>" align="right" style="border:thin solid #666666">	<% =XLS_vluDolaresUltimaCompra %></td>
							<td class="<%=CssUCTablaArt%>" align="center" style="border:thin solid #666666">	<% =XLS_idPIC %></td>
					<% end if %>
				</tr>
				<% Call resetArticuloVlu() %>
			<% rsStock.movenext %>		
		<% Wend %>
	</table>
	</body>
	</html>
		<%	
end if

%>