<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->

<%
Set oDiccCantidadesPedidas  = createObject("Scripting.Dictionary")
'-------------------------------------------------------------------------------------------------'
Function getCantidadPedida(pIdArticulo)
	Dim rtrn

	rtrn = 0
	if (oDiccCantidadesPedidas.Exists(cdbl(pIdArticulo))) then
		rtrn = oDiccCantidadesPedidas.Item(cdbl(pIdArticulo))
	end if

	getCantidadPedida = rtrn
End Function
'-------------------------------------------------------------------------------------------------'
Dim strSQL,idAlmacen,rs1,conn,stockEnPics,picIcon

	idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
	articulos =  GF_PARAMETROS7("articulos","",6)
	totalItems =  GF_PARAMETROS7("totalItems",0,6)

	strSQL =          "SELECT a.*, "
	strSQL = strSQL & "       ( existencia + sobrante ) stock, "
	strSQL = strSQL & "       art.dsarticulo "
	strSQL = strSQL & "FROM   (SELECT * "
	strSQL = strSQL & "        FROM   tblarticulosdatos "
	strSQL = strSQL & "        WHERE  idalmacen = " & idAlmacen
	strSQL = strSQL & "               AND ( existencia + sobrante ) < stockminimo "
	strSQL = strSQL & "               AND stockminimo <> 0) a "
	strSQL = strSQL & "       INNER JOIN tblarticulos art "
	strSQL = strSQL & "         ON a.idarticulo = art.idarticulo "
	strSQL = strSQL & "ORDER  BY A.idarticulo "
	call executeQueryDb(DBSITE_SQL_INTRA, rs1, "OPEN", strSQL)

	Set oDiccCantidadesPedidas = cargarCantidadesPedidas(idAlmacen,0)
%>
<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css">

		<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
	</head>
	<body>
		<table class="reg_header" width='100%'>
			<tr>
				<th colspan='3' class="reg_header_nav ui-corner-top">
					Faltante de articulos
				</th>
			</tr>

<%
	if (rs1.EoF) then 
	%>
		<tr>
			<th colspan='4' >
				Sin articulos
			</th>
		</tr>	
	<%
	end if
	while not rs1.EoF
		idarticulo = rs1("idarticulo")
		dsarticulo = rs1("dsarticulo")
		stockEnPics = getCantidadPedida(idarticulo)

		stockAPedir = cdbl(rs1("stockMinimo")) - stockEnPics - cdbl(rs1("stock"))
		if (stockAPedir < 0) then stockAPedir = 0

		if (stockEnPics <> 0) then
			'faltante de articulos pero hay pedidos'	
			if ((cdbl(rs1("stock")) + stockEnPics) < cdbl(rs1("stockMinimo"))) then
				'con el pedido aun no alcanza a cubrir el minimo'
				msgAlerta = "<img src='images/almacenes/alert-16x16.png' title='Stock Insuficiente' style='cursor:help'>"
				classAlerta = "reg_header_error"
			else
				'con el pedido alcanza a cubrir el minimo'
				stockAPedir = 0
				msgAlerta = "<img src='images/almacenes/warning-16x16.png' title='"&stockEnPics&" Articulo/s Pedido/s' style='cursor:help'>"
				classAlerta = "reg_header_warning"
			end if
		else
			msgAlerta = "<img src='images/almacenes/alert-16x16.png' title='Stock Insuficiente' style='cursor:help'>"
			classAlerta = "reg_header_error"
		end if



		%>
			<tr class="<%=classAlerta%>">
				<td>
					<input type="checkbox" name="articulo<%=totalItems%>" id="articulo<%=totalItems%>" value="<%=idarticulo&";"&stockAPedir%>">
				</td>
				<td>
					<label id="articulo_text_<%=totalItems%>"><%=idarticulo & " - " & dsarticulo%></label>
				</td>
				<td align='center' class='iconCell'>
					<%=msgAlerta%>
				</td>
			</tr>
		<%
		totalItems = totalItems + 1
		rs1.MoveNext
	wend

%>
		</table>
		<input type="hidden" id="totalitems" name="totalitems" value="<%=totalItems%>">
	</form>
	</body>
</html>



