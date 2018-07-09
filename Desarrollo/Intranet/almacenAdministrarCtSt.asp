<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
Const ESTADO_TODOS = -1
Const ESTADO_COMPLETO = 12
Const CANTIDAD_MINIMA_STOCK = 20
'-------------------------------------------------------------------------------
Function cargarControlStock(pIdControl,pCdResultado,pCdResponsable,pFecha,pAlmacen,pDivision,pEstado, pRsAlmacenes, pIdArticulo)
dim rs, strSQL,myWhere
call buscarFiltrosControlStock(myWhere,pIdControl,pCdResultado,pCdResponsable,pFecha,pAlmacen,pDivision,pEstado, pRsAlmacenes)
'Busqueda especial por artículo. SE hace así para seguiur devolviendo solo una cabecera.
strTablaCtrlStock = "tblcstkcabecera"
if (pIdArticulo > 0) then
    strTablaCtrlStock = "(Select Distinct C.* from " & strTablaCtrlStock & " C inner join TBLCSTKDETALLE D on C.IDCONTROL=D.IDCONTROL where D.IDARTICULO=" & pIdArticulo & ")"
end if
strSQL = " 	 	 SELECT A.idcontrol,"
strSQL = strSQL & " 	B.idalmacen,"
strSQL = strSQL & " 	B.dsalmacen,"
strSQL = strSQL & " 	A.momento,"
strSQL = strSQL & " 	A.cdresponsable,"
strSQL = strSQL & " 	D.idvale,"
strSQL = strSQL & " 	D.nrvale,"
strSQL = strSQL & " 	D.Fecha FechaResultado,"
strSQL = strSQL & " 	A.idResultado,"
strSQL = strSQL & " 	A.idEstado,"
strSQL = strSQL & " 	A.cantidad,"
strSQL = strSQL & " 	A.preciominimo,"
strSQL = strSQL & " 	A.preciomaximo,"
strSQL = strSQL & " 	A.artconstock,"
strSQL = strSQL & " 	A.tipo"
strSQL = strSQL & " FROM "
strSQL = strSQL & strTablaCtrlStock & " A "
strSQL = strSQL & " LEFT JOIN (SELECT E.idalmacen From  tblalmacenesusuario E "
strSQL = strSQL & " 		where E.cdusuario = '" & session("Usuario") & "' and E.nivel <> '"&ALMACEN_SOLICITANTE&"') as T "
strSQL = strSQL & " 	on T.idalmacen = A.idalmacen"
strSQL = strSQL & " LEFT JOIN tblalmacenes B"
strSQL = strSQL & " 	ON B.idalmacen = A.idalmacen"
strSQL = strSQL & " LEFT JOIN tbldivisiones C"
strSQL = strSQL & " 	ON C.iddivision = B.iddivision"
strSQL = strSQL & " LEFT JOIN (Select * from tblvalescabecera where ESTADO = " & ESTADO_ACTIVO & ") D"
strSQL = strSQL & " 	ON D.idvale = A.idresultado"
strSQL = strSQL &	 	myWhere
strSQL = strSQL & " ORDER  BY idcontrol DESC "
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
Set cargarControlStock = rs
End function
'-------------------------------------------------------------------------------
Function buscarFiltrosControlStock(ByRef myWhere,pIdControl,pCdResultado,pCdResponsable,pFecha,pAlmacen,pDivision,pEstado, pRsAlmacenes)	
	'Se agrega el filtro obligado de almacenes.
	myWhere = " WHERE B.IDALMACEN IN ("
	while not pRsAlmacenes.eof
	    myWhere = myWhere & pRsAlmacenes("IDALMACEN") & ", "
        pRsAlmacenes.MoveNext()
    wend	
    myWhere = Left(myWhere, Len(myWhere) - 2)
	myWhere = myWhere & ")"
	pRsAlmacenes.MoveFirst()
	'Se agregan los filtros de busqueda.
	if (pAlmacen > 0) then	Call mkWhere(myWhere, "B.IDALMACEN", pAlmacen, "=", 1)	
	if (pIdControl <> "") then Call mkWhere(myWhere, "A.IDCONTROL", pIdControl, "=", 1)
	if (pCdResultado <> "") then Call mkWhere(myWhere, "D.NRVALE", pCdResultado, "LIKE", 3)
	if (pFecha <> "") then Call mkWhere(myWhere, "A.MOMENTO", pFecha, "LIKE", 3)
	if ((pDivision <> "") and (pDivision <> 0)) then Call mkWhere(myWhere, "C.IDDIVISION", pDivision, "=", 1)	
	if (pCdResponsable <> "") then  Call mkWhere(myWhere, "A.CDRESPONSABLE", pCdResponsable, "LIKE", 3)			
	'Se consideran completos todos los controles creados y que ya se realizaron (con resultados cargados)
	if (pEstado = ESTADO_COMPLETO)then 
		Call mkWhere(myWhere, "A.IDRESULTADO", "0", "<>", 1)
		Call mkWhere(myWhere, "A.IDESTADO", ESTADO_ACTIVO, "=", 1)
	end if
	if (pEstado = ESTADO_BAJA)then Call mkWhere(myWhere, "A.IDESTADO", ESTADO_BAJA, "=", 1)
	'Se consideran activos todos los controles creados y que aún no se realizaron (sin resultados)
	if (pEstado = ESTADO_ACTIVO)then
		Call mkWhere(myWhere, "A.IDESTADO", ESTADO_ACTIVO, "=", 1)
		Call mkWhere(myWhere, "A.IDRESULTADO", "0", "=", 1)
	end if	
	buscarFiltrosControlStock = myWhere
End function
'-------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Dim fecActualD,fecActualM,fecActualA,IdControl,CodigoResultado,dsResponsable,cdResponsable
Dim IdAlmacen,IdDivision,cdEstado,rsAlmacenes,hayBusqueda,rsDivision,my_strSQL
Dim rsCnStk,mostrar,paginaActual,lineasTotales,reg,rsUser,busquedaActiva, idArticulo, dsArticulo

'Para poder acceder se debe tener permisos sobre el almacen (> solicitante).
Set rsAlmacenes = obtenerListaAlmacenesUsuario()
if rsAlmacenes.eof then	response.redirect "comprasAccesoDenegado.asp"

IdControl = GF_PARAMETROS7("IdControl", "", 6)
call addParam("IdControl", IdControl, params)
CodigoResultado = GF_PARAMETROS7("CodigoResultado", "", 6)
call addParam("CodigoResultado", CodigoResultado, params)
cdResponsable = GF_PARAMETROS7("cdResponsable", "", 6)
call addParam("cdResponsable", cdResponsable, params)
dsResponsable = getUserDescription(cdResponsable)
IdAlmacen = GF_PARAMETROS7("IdAlmacen", 0, 6)
call addParam("IdAlmacen", IdAlmacen, params)
idArticulo = GF_PARAMETROS7("idArticulo", 0, 6)
if (idArticulo <> 0) then Call getArticuloFull(idArticulo, dsArticulo, "")
call addParam("idArticulo", idArticulo, params)
IdDivision = GF_PARAMETROS7("IdDivision", 0, 6)
call addParam("IdDivision", IdDivision, params)
cdEstado = GF_PARAMETROS7("cdEstado", 0, 6)
if (cdEstado = 0) then cdEstado = ESTADO_ACTIVO
call addParam("cdEstado", cdEstado, params)
fecActualD = GF_PARAMETROS7("fecActualD", "", 6)
if (fecActualD <> "") then fecActualD = GF_nDigits(fecActualD,2)
call addParam("fecActualD", fecActualD, params)
fecActualM = GF_PARAMETROS7("fecActualM", "", 6)
if (fecActualM <> "") then fecActualM = GF_nDigits(fecActualM,2)
call addParam("fecActualM", fecActualM, params)
fecActualA = GF_PARAMETROS7("fecActualA", "", 6)
if (fecActualA <> "") then fecActualA = GF_nDigits(fecActualA,4)
call addParam("fecActualA", fecActualA, params)
My_Fecha = Trim(fecActualA&fecActualM & fecActualD)
hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", busquedaActiva, params)
if busquedaActiva = 1 then hayBusqueda = true

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

Set rsCnStk = cargarControlStock(IdControl,CodigoResultado,cdResponsable,My_Fecha,IdAlmacen,IdDivision,cdEstado, rsAlmacenes, idArticulo)
Call setupPaginacion(rsCnStk, paginaActual, mostrar)
lineasTotales = rsCnStk.recordcount

%>
<html>
<head>
<title><%=GF_TRADUCIR("Administracion Control de Stock")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.divOculto {
	display: none;
}
</style>
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">	
	var ch = new channel();
	
	function bodyOnLoad() {
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("Previous-16x16.png", "Volver", "Volver()");
		tb.addButton("refresh-16x16.png", "Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.addButton("add-16x16.png", "Nuevo Control Stock", "loadPopUpNew()");				
		tb.draw();
		<%	if (hayBusqueda) then %>
				tb.changeState(swt);
				startMagicSearch();
		<%	end if 
		 	if (not rsCnStk.eof) then %>
				var pgn = new Paginacion("paginacion");
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "almacenAdministrarCtSt.asp<% =params %>");					
		<%	end if 	%>
		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;
		msArticulo.setValue('<% =dsArticulo %>')		
	}
	
	function buscarOn() {
		document.getElementById("busqueda").className = "";	
		document.getElementById("busquedaActiva").value = "1";			
		startMagicSearch();
	}
	
	function buscarOff() {
		document.getElementById("busqueda").className = "divOculto";		
		document.getElementById("busquedaActiva").value = "0";		
	}
	
	function EliminarControl(pIdControl){
		ch.bind("almacenCtSt_Ajax.asp?accion=<%=ACCION_BORRAR%>&IdControl="+ pIdControl ,"loadPopUpNew_callback()");
		ch.send();
	}
	
	function ReactivarControl(pIdControl){
		ch.bind("almacenCtSt_Ajax.asp?accion=<%=ACCION_ACTIVAR%>&IdControl="+ pIdControl ,"loadPopUpNew_callback()");
		ch.send();
	}
	
	function loadPopUpNew_callback(){
		submitInfo();
	}
		
	
	function Volver(){
		location.href = "almacenAuditoria.asp";
	}
	function submitInfo() {
		document.getElementById("frmSel").submit();
	}

	function loadPopUpNew(id) {		
		puw = new winPopUp('popupNuevoCtSt','almacenPropCtStNuevo.asp?cantArticulostxt=<%=CANTIDAD_MINIMA_STOCK%>', '420','240','Nuevo Control de Stock', "loadPopUpNew_callback()");
	}		
		
	function startMagicSearch() {			
		var msResponsable = new MagicSearch("", "divResponsable", 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msResponsable.setToken(";");
		msResponsable.onBlur = seleccionarResponsable;
		msResponsable.setValue('<%=dsResponsable%>');
	}
	function seleccionarResponsable(ms) {				
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdResponsable").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdResponsable").value = "";
		}		
	}
	function imprimirControlStock(pIdControl,pIdAlmacen,pcdResponsable,pDsAlmacen,pIdResultado){
		var myCantArt = document.getElementById("cantidadArt_" + pIdControl).value;		
		var myPrecioMinimo = document.getElementById("precioMinimo_" + pIdControl).value;
		var myArtconStock = document.getElementById("artconStock_" + pIdControl).value;
		var myTipoRep = document.getElementById("tipoReporte_" + pIdControl).value;
		var myPrecioMaximo = document.getElementById("precioMaximo_" + pIdControl ).value;
		var ventana = window.open("almacenCtStPrint.asp?idControl="+pIdControl+"&idAlmacen="+pIdAlmacen+"&cdResponsable="+pcdResponsable+"&dsAlmacen="+pDsAlmacen+"&IdResultado="+pIdResultado+"&cantArt="+myCantArt+"&precioMaximo="+myPrecioMaximo+"&precioMinimo="+myPrecioMinimo+"&tipoReporte="+myTipoRep+"&artconStock="+myArtconStock, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	function cargarResultados(pIdControl,pIdAlmacen,pCdResponsable){		
		var myTipoRep = document.getElementById("tipoReporte_" + pIdControl).value;
		var respuesta = window.showModalDialog("almacenCtStSelecciones.asp?IdAlmacen="+pIdAlmacen+"&IdControl="+pIdControl+"&accion=<%=ACCION_PROCESAR%>&tipoReporte="+myTipoRep+"&cdResponsable="+pCdResponsable,"_blank","dialogHeight:500px;dialogLeft=400px;dialogWidth:700px;center=yes;scroll:yes")
		if(respuesta == true) submitInfo();		
	}
	function abrirVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function seleccionarArticulo(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("idArticulo").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idArticulo").value = "";							
		}		
	}
</script>
</head>
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Control de Stock") %>
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="almacenAdministrarCtSt.asp" method="POST">
<div id="busqueda" class="divOculto">
	<br><br>	
	<table id="tblBusqueda" width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
           <td width="8"><img src="images/marco_r2_c3.gif"></td>
           <td align="right"></td>
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
                     <table width="95%" align="center" border="0">
                            <tr>								
								<td width="15%" align="right"><% = GF_TRADUCIR("Id Control") %>:</td>
								<td width="20%">
									<input type="text"  id="IdControl" name="IdControl" value="<%=IdControl%>">
								</td>
								<td width="13%" align="right"><% = GF_TRADUCIR("Codigo Resultado") %>:</td>
								<td width="20%">
									<input type="text" size="15" id="CodigoResultado" name="CodigoResultado" value="<%=CodigoResultado%>">									
								</td>
                            </tr>     
							<tr>								
								<td align="right"><% = GF_TRADUCIR("Responsable") %>:</td>
								<td>
									<div id="divResponsable"></div>			
									<% =dsSolicitante %>
									<input type="hidden" id="cdResponsable" name="cdResponsable" value="<% =cdResponsable %>">
								</td>
								<td width="35%" align="right"><% = GF_TRADUCIR("Fecha") %>:</td>
								<td>
									<input type="text" size="1" maxLength="2" value="<% =fecActualD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualD" id="fecActualD"> /
									<input type="text" size="1" maxLength="2" value="<% =fecActualM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualM" id="fecActualM"> /
									<input type="text" size="2" maxLength="4" value="<% =fecActualA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecActualA" id="fecActualA">			
								</td>	
                            </tr>     
							<tr>														
							
								<td align="right"><% =GF_TRADUCIR("Almacen") %>:</td>
                                <td>                                
									<select id="idAlmacen" name="idAlmacen">
											<option value="0" <% if (IdAlmacen=0) then response.write "selected='true'" %>><% =GF_TRADUCIR("Todos") %>
											<%	
											while (not rsAlmacenes.eof)	
											%>
												<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = IdAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
											<%		
												rsAlmacenes.MoveNext()
											wend 	
											%>		
									</select>		
                                </td>	
								<td align="right"><% = GF_TRADUCIR("Division")     %>:</td>
								<td>
									<%
									my_strSQL="Select * from TBLDIVISIONES"
									call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", my_strSQL)	
									%>
										<select id="idDivision" name="idDivision">
											<option value="" <%if (IdDivision = "") then %> selected='true' <%end if%>><% =GF_TRADUCIR("-Seleccione-") %></option>
											<%
											while (not rsDivision.eof)									
													 %>
														<option value="<% =rsDivision("IDDIVISION") %>" <% if (IdDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
											<%														
												rsDivision.MoveNext()
											wend	
											%>								
										</select>							
								</td>
							</tr>	
							<tr>
								<td align="right"><% = GF_TRADUCIR("Estado")%>:</td>
								<td>							
									<select id="cdEstado" name="cdEstado">
										<option value="<%=ESTADO_TODOS%> "	  <%if (cdEstado = ESTADO_TODOS)     then response.write "selected='true'"%> ><%=GF_TRADUCIR("-Todos-")%>
										<option value="<%=ESTADO_ACTIVO%>"	  <%if (cdEstado = ESTADO_ACTIVO)	 then response.write "selected='true'" %>><%=GF_TRADUCIR("Pendientes")%>
										<option value="<%=ESTADO_COMPLETO%>"  <%if (cdEstado = ESTADO_COMPLETO)  then response.write "selected='true'" %>><%=GF_TRADUCIR("Realizado")%>
										<option value="<%=ESTADO_BAJA%>"	  <%if (cdEstado = ESTADO_BAJA)		 then response.write "selected='true'" %>><%=GF_TRADUCIR("Anulado")%>
									</select>							
								</td>
								<td align="right"><% =GF_TRADUCIR("Articulo") %></td>
	                            <td >
		                            <div id="articuloItem0"></div>																		
		                            <input type="hidden" id="idArticulo" name="idArticulo" value="<% =idArticulo %>">
	                            </td>
                            </tr>	
                            <tr>															
	                            <td colspan="4"  align="center"><input type="submit" value="Buscar" id="submit1" name="submit1" onclick="submitInfo();"></td>
                            </tr>								
                     </table>
	           </td>
	           <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="images/marco_r3_c1.gif"></td>
	           <td width="100%" align="center" colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="images/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	</div> 	
	<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">
	<input type="hidden" name="accion" id="accion" value="<%=ACCION_SUBMIT%>">		
	
<br>
	
	<br>
<table class="reg_Header" align="center" width="70%" border="0">

	<tr><td colspan="6"><% Call showErrors() %></td></tr>	
	<tr><td colspan="6"><div id="paginacion"></div></td></tr>				
	<tr>
		<td class="reg_header_nav"  align="center"><%=GF_Traducir("Id")%></td>
		<td class="reg_header_nav"  align="center"><%=GF_Traducir("Almacen")%></td>
		<td class="reg_header_nav"  align="center"><%=GF_Traducir("Fecha")%></td>
		<td class="reg_header_nav"  align="center"><%=GF_Traducir("Responsable")%></td>
		<td width="32px" class="reg_header_nav"  align="center"><%=GF_Traducir("Reporte")%></td>		
		<td class="reg_header_nav" colspan="3" align="center"><%=GF_Traducir("Resultado")%></td>
		<td width="32px" class="reg_header_nav"  align="center"><%=GF_Traducir(".")%></td>
	</tr>
	<% 
	if rsCnStk.eof then %>
		<tr>
			<td align="center" colspan="8">
				<%=GF_Traducir("No se encontraron resultados")%>
			</td>
		</tr>
	<% end if
		while ((not rsCnStk.eof) and (CInt(reg) < CInt(mostrar)))
			reg = reg + 1%>
			<tr class="reg_Header_navdos <%if (CInt(rsCnStk("IDESTADO")) = ESTADO_BAJA) then Response.write "reg_header_rejected" %>">
				<td  align="center"><%=rsCnStk("IDCONTROL")%></td>
				<td  align="center"><%=rsCnStk("DSALMACEN")%></td>
				<td  align="center"><%=GF_FN2DTE(left(rsCnStk("MOMENTO"),8))%></td>
				<td  align="center"><%=rsCnStk("CDRESPONSABLE")&"-"&getUserDescription(rsCnStk("CDRESPONSABLE"))%></td>
				<td  align="center" onclick="imprimirControlStock(<%=rsCnStk("IDCONTROL")%>,<%=rsCnStk("IDALMACEN")%>,'<%=rsCnStk("CDRESPONSABLE")%>','<%=rsCnStk("DSALMACEN")%>',<%=rsCnStk("IDRESULTADO")%>)"><img title="Imprimir Control Stock" src="images/almacenes/printer-16x16.png" style="cursor:pointer"></td>				
				<% if (rsCnStk("IDRESULTADO") <> CTST_SIN_VALE) then %>
				    <td  align="center"><%=GF_FN2DTE(rsCnStk("FECHARESULTADO"))%></td>
				    <td  align="center"><%=rsCnStk("NRVALE")%></td>
				<% else %>				
				    <td  align="center" colspan="2">- SIN DIFERENCIAS -</td>
				<% end if
				  if(rsCnStk("IDVALE")>0)then%>
					<td width="15px"  align="center" onclick="abrirVale(<%=rsCnStk("IDVALE")%>)"><img title="Imprimir Vale" src="images/almacenes/printer-16x16.png" style="cursor:pointer"></td>
					<td width="15px"  align="center" ></td>
				<%elseif (CInt(rsCnStk("IDESTADO")) = ESTADO_BAJA) then	%>			
					<td width="15px"  align="center" ></td>		
					<td width="15px"  align="center" onclick="ReactivarControl(<%=rsCnStk("IDCONTROL")%>)"><img title="Reactivar" src="images/compras/accept-16x16.png" style="cursor:pointer"></td>					
				<%elseif (rsCnStk("IDRESULTADO") = CTST_SIN_VALE) then %>
				    <td width="15px"  align="center" ></td>
				    <td width="15px"  align="center" ></td>
				<%else %>	
					<td width="15px"  align="center" onclick="cargarResultados(<%=rsCnStk("IDCONTROL")%>,<%=rsCnStk("IDALMACEN")%>,'<%=rsCnStk("CDRESPONSABLE")%>')"><img title="Cargar Resultado Control" src="images/almacenes/edit-16x16.png" style="cursor:pointer"></td>				
					<td width="15px"  align="center" onclick="EliminarControl(<%=rsCnStk("IDCONTROL")%>)"><img title="Eliminar" src="images/compras/cancel-16x16.png" style="cursor:pointer"></td>				
				<%end if%>
				<input type="hidden" id="cantidadArt_<%=rsCnStk("IDCONTROL")%>" name="cantidadArt_<%=rsCnStk("IDCONTROL")%>" value="<%=rsCnStk("CANTIDAD")%>">
				<input type="hidden" id="precioMinimo_<%=rsCnStk("IDCONTROL")%>" name="precioMinimo_<%=rsCnStk("IDCONTROL")%>" value="<%=rsCnStk("PRECIOMINIMO")%>"> 
				<input type="hidden" id="precioMaximo_<%=rsCnStk("IDCONTROL")%>" name="precioMaximo_<%=rsCnStk("IDCONTROL")%>" value="<%=rsCnStk("PRECIOMAXIMO")%>"> 				
				<input type="hidden" id="artconStock_<%=rsCnStk("IDCONTROL")%>" name="artconStock_<%=rsCnStk("IDCONTROL")%>" value="<%=rsCnStk("ARTCONSTOCK")%>">
				<input type="hidden" id="tipoReporte_<%=rsCnStk("IDCONTROL")%>" name="tipoReporte_<%=rsCnStk("IDCONTROL")%>" value="<%=rsCnStk("TIPO")%>">
			</tr>
		<%rsCnStk.movenext
		wend %>
</table>
</form>

</body>
</html>