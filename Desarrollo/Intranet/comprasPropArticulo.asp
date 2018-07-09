<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Function controlar(idArticulo, dsArticulo, idCategoria, idUnidad, stMinimo, stMaximo, cmMinima, cmMaxima, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
	Dim strSQL, rs, conn
	
	controlar = RESPUESTA_OK
	'Limites de Stock
	if (controlar = RESPUESTA_OK) then
		if ((stMinimo <> 0) or (stMaximo <> 0))then
			'Se definieron limites de stock.
			if (CDbl(stMinimo) > CDbl(stMaximo)) then
				controlar = LIMITES_STOCK
			end if
		end if
	end if
	'Limites de Compra
	if (controlar = RESPUESTA_OK) then
		if ((cmMinima <> 0) or (cmMaxima <> 0))then
			'Se definieron limites de compra
			if (CDbl(cmMinima) > CDbl(cmMaxima)) then
				controlar = LIMITES_COMPRA
			end if
		end if
	end if	
End Function
'-----------------------------------------------------------------------------------------
Function accionGrabar(idArticulo, idAlmacen, cdInterno, dsArticulo, idCategoria, idUnidad, stMinimo, stMaximo, cmMinima, cmMaxima, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
	Dim strSQL, rs, conn
	
	'Solo cuando creo articulos o los modifico desde compras.
	if (idAlmacen = 0) then Call grabarArticulo(idArticulo, dsArticulo, idCategoria, idUnidad, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
	'Solo cuando modifico desde el almacen.
	if (idAlmacen <> 0) then Call grabarDatosArticulo(idArticulo, idAlmacen, cdInterno, stMinimo, stMaximo, cmMinima, cmMaxima)
	accionGrabar = true	
End Function
'-----------------------------------------------------------------------------------------
Sub grabarArticulo(idArticulo, dsArticulo, idCategoria, idUnidad, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
	Dim strSQL, rs, conn
	
	if (idArticulo = 0) then
		'Es una unidad nueva
		strSQL="Insert into TBLARTICULOS(DSARTICULO, IDCATEGORIA, IDUNIDAD, CDCUENTA, BIENUSO, ESTADO, CDCUENTAGASTOS, CCOSTOS, CDCUENTASAF, CDUSUARIO, MOMENTO, IDREEMPLAZO)"
		strSQL= strSQL & " values('" & UCase(dsArticulo) & "', " & idCategoria & ", " & idUnidad & ",'" & cdCuenta & "', '" & bienUso & "', " & ESTADO_ACTIVO & ", '" & cdCuentaGastos & "','" & cCosto & "', '" & cdCuentaSAF & "','" & session("Usuario") & "', " & session("MmtoSistema") & ", 0)"
	else
		'Es una modificacion
		strSQL="Update TBLARTICULOS Set DSARTICULO='" & Ucase(dsArticulo) & "', CDCUENTA='" & cdCuenta & "', IDCATEGORIA=" & idCategoria
		strSQL = strSQL & ", IDUNIDAD=" & idUnidad & ", BIENUSO='" & bienUso & "', CDCUENTAGASTOS='" & cdCuentaGastos & "', CCOSTOS='" & cCosto & "'"
		strSQL = strSQL & ", CDCUENTASAF='" & cdCuentaSAF & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema")		
		strSQL = strSQL & " where IDARTICULO=" & idArticulo		
	end if
	'response.write strSQL		
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Sub
'-----------------------------------------------------------------------------------------
Sub grabarDatosArticulo(idArticulo, idAlmacen, cdInterno, stMinimo, stMaximo, cmMinima, cmMaxima)
	Dim strSQL, rs, conn
	Dim l_idArticulo
	if (idArticulo = 0) then
		'traer articulo que recien se ha grabado en tblarticulos
		strSQL = "select max(idarticulo) as idArticuloNuevo from tblarticulos "
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then l_idArticulo = rs("idArticuloNuevo")
	else
		l_idArticulo = idArticulo
	end if
	'chequear, si en la tabla tblarticulosdatos ya existe articulo en el almacen dado - decidir, si se inserta nuevo registro o si se actualiza el existente
	strSQL = "select * from tblarticulosdatos where idarticulo = " & l_idArticulo & " and idalmacen = " & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if rs.eof then
		'se inserta nuevo registro
		strSQL="Insert into TBLARTICULOSDATOS(IDARTICULO, IDALMACEN, EXISTENCIA, SOBRANTE, STOCKMINIMO, STOCKMAXIMO, COMPRAMINIMA, COMPRAMAXIMA, CDINTERNO, CDUSUARIO, MOMENTO)"
		strSQL= strSQL & " values(" & l_idArticulo & ", " & idAlmacen & ", 0, 0, " & stMinimo & ", " & stMaximo & ", " & cmMinima & ", " & cmMaxima & ", '" & cdinterno &  "', '" & session("Usuario") & "', " & session("MmtoSistema") & ")"				
	else
		'Es una modificacion
		strSQL="Update TBLARTICULOSDATOS Set "
		strSQL = strSQL & " STOCKMINIMO=" & stMinimo & ", STOCKMAXIMO=" & stMaximo & ", COMPRAMINIMA=" & cmMinima & ", COMPRAMAXIMA=" & cmMaxima
		strSQL = strSQL & ", CDINTERNO='" & cdInterno & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema")		
		strSQL = strSQL & " where IDARTICULO=" & idArticulo & " and idalmacen= " & idAlmacen		
	end if	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
End Sub
'-----------------------------------------------------------------------------------------
Function accionConsulta(idArticulo, ByRef dsArticulo, ByRef idCategoria, ByRef idUnidad, ByRef cdCuenta, ByRef bienUso, byRef cdCuentaGastos, byRef cCosto, byRef cdCuentaSAF, ByRef dsCategoria, ByRef dsUnidad, ByRef idReemplazo)
	Dim strSQL, rs, conn	
	strSQL="Select * from TBLARTICULOS where IDARTICULO=" & idArticulo
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then		
		dsArticulo = Trim(rs("DSARTICULO"))
		cdCuenta = Trim(rs("CDCUENTA"))
		idCategoria = rs("IDCATEGORIA")
		idUnidad = rs("IDUNIDAD")
		bienUso = Trim(rs("BIENUSO"))
		cdCuentaGastos = rs("CDCUENTAGASTOS")
		cdCuentaSAF = rs("CDCUENTASAF")
		cCosto = Trim(rs("CCOSTOS"))
		idReemplazo = Trim(rs("IDREEMPLAZO"))
		strSQL="Select * from TBLARTCATEGORIAS where IDCATEGORIA=" & rs("IDCATEGORIA")
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
		if (not rs2.eof) then dsCategoria = rs2("DSCATEGORIA")
		strSQL="Select * from TBLUNIDADES where IDUNIDAD=" & rs("IDUNIDAD")
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
		if (not rs2.eof) then dsUnidad = rs2("DSUNIDAD")				
	end if	
End Function
'-----------------------------------------------------------------------------------------
Function accionConsultaDatos(idArticulo, idAlmacen, ByRef stMinimo, ByRef stMaximo, ByRef cmMinima, ByRef cmMaxima, ByRef cdInterno)
	Dim strSQL, rs, conn	
	strSQL="Select * from TBLARTICULOSDATOS where IDARTICULO=" & idArticulo & " and idalmacen = " & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then		
		cdInterno = Trim(rs("CDINTERNO"))
		stMinimo = rs("STOCKMINIMO")
		stMaximo = rs("STOCKMAXIMO")
		cmMinima = rs("COMPRAMINIMA")
		cmMaxima = rs("COMPRAMAXIMA")	
	end if	
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idArticulo, dsArticulo, idCategoria, idUnidad, stMinimo, stMaximo, cmMinima, cmMaxima, cdCuenta, bienUso, cdInterno, cdCuentaGastos, cdCuentaSAF, cCosto, idReemplazo


idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
'Si entra por compras, se controla que tenga permiso.
if (idAlmacen = 0) then Call comprasControlAccesoCM(RES_ADM)

idArticulo = GF_PARAMETROS7("idArticulo",0,6)

cdInterno = GF_PARAMETROS7("cdInterno","",6)
dsArticulo = UCase(GF_PARAMETROS7("descripcion","",6))
cdCuenta = GF_PARAMETROS7("cuenta","",6)
cdCuentaGastos = GF_PARAMETROS7("cuentaGastos","",6)
cdCuentaSAF = GF_PARAMETROS7("cuentaSAF","",6)
cCosto = GF_PARAMETROS7("cCosto","",6)
idCategoria = GF_PARAMETROS7("idCategoria",0,6)
idUnidad = GF_PARAMETROS7("idUnidad",0,6)
stMinimo = GF_PARAMETROS7("stockMinimo",2,6)
stMaximo = GF_PARAMETROS7("stockMaximo",2,6)
cmMinima = GF_PARAMETROS7("compraMinima",2,6)
cmMaxima = GF_PARAMETROS7("compraMaxima",2,6)
bienUso = GF_PARAMETROS7("bienUso","",6)
accion = GF_PARAMETROS7("accion","",6)

Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	errMsg = controlar(idArticulo, dsArticulo, idCategoria, idUnidad, stMinimo, stMaximo, cmMinima, cmMaxima, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
	if (errMsg = RESPUESTA_OK) then
		Call accionGrabar(idArticulo, idAlmacen, cdInterno, dsArticulo, idCategoria, idUnidad, stMinimo, stMaximo, cmMinima, cmMaxima, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF)
		accion = ACCION_CERRAR
	else
		setError(errMsg)
	end if
else
	Call accionConsulta(idArticulo, dsArticulo, idCategoria, idUnidad, cdCuenta, bienUso, cdCuentaGastos, cCosto, cdCuentaSAF, dsCategoria, dsUnidad, idReemplazo)
	if idAlmacen <> 0 then Call accionConsultaDatos(idArticulo, idAlmacen, stMinimo, stMaximo, cmMinima, cmMaxima, cdInterno)	
end if
if (accion = "") then accion = ACCION_GRABAR
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
var refPopUpArticulo;
var ch = new channel();
 	
function descripcionOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function cargarCombo(comboName, ls, elemId, showIndex) {
	var combo = document.getElementById(comboName);
	//Borro todo.
	for(; combo.length > 0;) combo.remove(0);
	//Se rearma la lista de unidades
	for(var i in ls) {		
		var opt=document.createElement('option');
		opt.value = ls[i][0];
 		opt.text = limitarString(ls[i][showIndex],50);
 		opt.title = ls[i][showIndex];
 		if (ls[i][0] == elemId) opt.selected = true;
		try {
    		combo.add(opt, null); // standards compliant
    	} catch(ex)	{
    		combo.add(opt); // IE only
    	}
	}
}

function agregarUnidad_callback() {
	var resp = ch.parsedResponse(false, ";", "|");	 	
	cargarCombo("idUnidad", resp, <% =idUnidad %>, 1);
}

function  agregarCategoria_callback() {	
	var resp = ch.parsedResponse(false, ";", "|");
	cargarCombo("idCategoria", resp, <% =idCategoria %>, 2);
}

function agregarUnidad() {
	refPopUpArticulo.resize(500, 420);
	ch.bind("comprasStreamElementos.asp?tipo=unidades", "agregarUnidad_callback");
	ch.send();
	return true;
}

function agregarCategoria() {
	refPopUpArticulo.resize(500, 420);
	ch.bind("comprasStreamElementos.asp?tipo=categorias", "agregarCategoria_callback");
	ch.send();
	return true;
}

function nuevaUnidad() {
	refPopUpArticulo.resize(500, 420);
	var puw = new winPopUp('popupUnidad', 'comprasPropUnidad.asp', 450, 320, "Propiedades de Unidad", 'agregarUnidad()');
}

function nuevaCategoria() {
	refPopUpArticulo.resize(500, 420);
	var puw = new winPopUp('popupCategoria', 'comprasPropCategoria.asp', 450, 300, "Propiedades de Categoria", 'agregarCategoria()');
}

function articuloOnLoad() {		
	refPopUpArticulo = getObjPopUp('popupArticulo');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpArticulo.hide();
	<% end if
	   if ((idAlmacen = 0) and (idArticulo = 0)) then %>
	//Se cargan las Categorias
	agregarCategoria();
	//Se cargan las unidades.
	agregarUnidad();	
	//Se enfoca en el primer campo.
	document.getElementById("descripcion").focus();	
	<% end if %>
}
</script>
</head>
<body onLoad="articuloOnLoad()">
<form name="frmSel" method="post" action="comprasPropArticulo.asp">
<table align="center">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/items-32x32.png"> <% =GF_TRADUCIR("Propiedades de Articulo") %></td>
	</tr>	
	<tr>
		<td colspan="2"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table border="0">				
		<%if (idAlmacen = 0) then %>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Codigo") %></td>
					<td> 
						<b><% =idArticulo %></b> 						
					</td>
					<td></td>
					<td></td>
				</tr>							
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td colspan="3"><input type="text" id="descripcion" name="descripcion" maxlength="50" size="50" value="<% =dsArticulo %>" onkeypress="return controlSalto(this, event)" onBlur="descripcionOnBlur(this)"></td>
				</tr>
			<%	if (idArticulo = 0) then %>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Categoria") %></td>
					<td colspan="3">
						<select id="idCategoria" name="idCategoria"></select>
						<a onClick="nuevaCategoria();"><img align="absMiddle" src="images/compras/categories_new-16x16.png" title="Agregar una categoria" style="cursor: pointer"></a>
					</td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Unidades") %></td>
					<td colspan="3">
						<select id="idUnidad" name="idUnidad"></select>
						<a onClick="nuevaUnidad();"><img align="absMiddle" src="images/compras/units_new-16x16.png" title="Agregar una unidad" style="cursor: pointer"></a>
					</td>					
				</tr>
			<%	else	%>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Categoria") %></td>
					<td colspan="3">
						<% =dsCategoria %>
						<input type="hidden" id="idCategoria" name="idCategoria" value="<% =idCategoria %>">
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Unidades") %></td>
					<td colspan="3">
						<% =dsUnidad %>	
						<input type="hidden" id="idUnidad" name="idUnidad" value="<% =idUnidad %>">
					</td>
				</tr>
			<%	end if	%>
				<tr>					
					<td class="reg_header"><% =GF_TRADUCIR("Cuenta") %></td>
					<td><input type="text" id="cuenta" name="cuenta" maxlength="9" size="10" value="<% =cdCuenta %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
					<td></td>
					<td></td>
				</tr>
				<tr>					
					<td class="reg_header"><% =GF_TRADUCIR("Tipo Bien") %></td>
					<td width="10%">
						<Select id="bienUso" name="bienUso">							
							<option value="<% =ES_BIEN_DE_CONSUMO %>" selected><% =GF_TRADUCIR("Bien de Consumo") %></option>
							<option value="<% =ES_BIEN_DE_USO %>" <% if (bienUso = ES_BIEN_DE_USO) then response.write "selected" %>><% =GF_TRADUCIR("Bien de Uso") %></option>
						</select>
					</td>
					<td></td>
					<td></td>					
				</TR>
				<tr>					
					<td class="reg_header"><% =GF_TRADUCIR("Cuenta Gastos") %></td>
					<td><input type="text" id="cuentaGastos" name="cuentaGastos" maxlength="12" size="10" value="<% =cdCuentaGastos %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
					<td class="reg_header"><% =GF_TRADUCIR("CC") %></td>
					<td><input type="text" id="cCosto" name="cCosto" maxlength="6" size="6" value="<% =cCosto  %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
				</tr>				
				<tr>					
					<td class="reg_header"><% =GF_TRADUCIR("Cuenta SAF") %></td>
					<td><input type="text" id="cuentaSAF" name="cuentaSAF" maxlength="12" size="11" value="<% =cdCuentaSAF%>" onKeyPress="return controlDatos(this, event, 'N')"></td>
				</tr>				
	<% else	
				Dim rsAlmacen
				Set rsAlmacen = obtenerListaAlmacenes(idAlmacen)
				%>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Codigo") %></td>
					<td> 
						<b><% =idArticulo %></b> 						
					</td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Categoria") %></td>
					<td colspan="3"><% =dsCategoria %></td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td colspan="3"><% =dsArticulo %></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Unidades") %></td>
					<td colspan="3"><% =dsUnidad %>	</td>					
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Almacen") %></td>
					<td colspan="3"><%=rsAlmacen("dsAlmacen")%></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Codigo Interno") %></td>
					<td colspan="3"><input type="text" id="cdInterno" name="cdInterno" maxlength="50" size="50" value="<% =cdInterno %>" onkeypress="return controlSalto(this, event)"></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Stock Min.") %></td>
					<td><input type="text" id="stockMinimo" name="stockMinimo" maxlength="10" size="10" value="<% =stMinimo %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
					<td class="reg_header"><% =GF_TRADUCIR("Stock Max.") %></td>
					<td><input type="text" id="stockMaximo" name="stockMaximo" maxlength="10" size="10" value="<% = stMaximo %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Compra Min.") %></td>
					<td><input type="text" id="compraMinima" name="compraMinima" maxlength="10" size="10" value="<% =cmMinima %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
					<td class="reg_header"><% =GF_TRADUCIR("Compra Max.") %></td>
					<td><input type="text" id="compraMaxima" name="compraMaxima" maxlength="10" size="10" value="<% =cmMaxima %>" onKeyPress="return controlDatos(this, event, 'N')"></td>
				</tr>
	<% end if	%>
			</table>
		</td>
	</tr>		
	<tr><td>&nbsp;</td><tr>
	<tr>
		<td></td>
		<td align="center">
			<table>	
				<tr><td>
					<%  if (not isAuditor(SIN_DIVISION)) then %>
					<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idArticulo = 0) then response.write "disabled=true" %>>					
					<%	end if	%>
				</td></tr>
			</table>
		</td>		
	</tr>	
</table>
<input type="hidden" id="idArticulo" name="idArticulo" value="<% =idArticulo %>">
<input type="hidden" name="idAlmacen" value="<% =idAlmacen %>">
<input type="hidden" name="accion" value="<% =accion %>">
</form>
</body>
</html>