<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<% 
Call controlAccesoAL(RES_ACC_AL)

Dim idAlmacen, rs, titleAux, flagUno, idArticulo, dsArticulo, cdCategoria, cdObra, idObra, idSector, strLink, flagLink, dsCategoria
Dim txtMesCierre, txtAnioCierre, txtFechaCierre, estadoCierre, tipoCuenta, cdCuenta, idVale, idDivision, idCierre, idBudgetArea, idBudgetDetalle, dsSector
idDivision = GF_Parametros7("idDivision", 0, 6)
txtMesCierre = GF_Parametros7("txtMesCierre", "", 6)
txtAnioCierre = GF_Parametros7("txtAnioCierre", "", 6)
idCierre = GF_PARAMETROS7("idCierre",0,6)
txtFechaCierre = txtAnioCierre & txtMesCierre & LastDayOfMonth(txtAnioCierre, txtMesCierre)
cdCuenta = trim(GF_Parametros7("cdCuenta", "", 6))
cdObra = GF_Parametros7("cdObra", "", 6)
idObra = GF_PARAMETROS7("idObra",0,6)
idSector = GF_PARAMETROS7("idSector",0,6)
idBudgetArea = GF_Parametros7("idBudgetArea", 0, 6)
idBudgetDetalle = GF_Parametros7("idBudgetDetalle", 0, 6)
idArticulo = GF_PARAMETROS7("idArticulo",0,6)
cdCategoria = GF_PARAMETROS7("cdCategoria","",6)
idCategoria = GF_PARAMETROS7("idCategoria",0,6)
estadoCierre = GF_PARAMETROS7("estadoCierre","",6)
tipoCuenta = GF_PARAMETROS7("tipoCuenta",0,6)
if idArticulo <> 0 then call getArticuloFull(idArticulo, dsArticulo, "")
if idSector <> 0 then dsSector = getSectorDS(idSector)
if idCategoria <> 0 then dsCategoria = getCategoriaDS(idCategoria)
'Filtrar division segun almacenes
'if idDivision = 0 then idDivision = 2
idAlmacen = getAlmacenesPorDivision(idDivision)
%>
<html>
<head>
<title>Administración Almacenes - Consultas Contables</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<style type="text/css">
.link {
	cursor:pointer;
	color:blue;
	text-decoration:underline;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	var ch = new channel();	
	var link;
	var hkey1;
	var hkey2;
	function irA(pLink) {
		location.href = pLink;
	}
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}	
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.addButton("refresh-16x16.png", "Refresh", "submitPage()");
		tb.addButton("Contabilidad_Folder-16x16.png", "Contabilidad", "irA('almacenCCN_Contabilidad.asp')");
		tb.draw();		
		realizarConsulta('<%=idDivision%>', '<%=idAlmacen%>', '<%=idCierre%>', '<%=estadoCierre%>', '<%=tipoCuenta%>', '<%=txtFechaCierre%>', '<%=cdCuenta%>', '<%=idObra%>', '<%=idBudgetArea%>', '<%=idBudgetDetalle%>', '<%=idCategoria%>', '<%=idArticulo%>', '<%=idSector%>');
	}
	function submitPage(){
		document.getElementById("frmSel").submit();
	}
	
	function realizarConsulta(division, almacen, idCierre, estadoCierre, tipoCuenta, fecha, cdCuenta, idObra, idBudgetArea, idBudgetDetalle, idCategoria, idArticulo, idSector){
		document.getElementById("imgLoading").style.position = "relative";
		document.getElementById("imgLoading").style.visibility  = "visible";
		document.getElementById("lblLoading").style.position = "relative";
		document.getElementById("lblLoading").style.visibility  = "visible";
		ch.bind("almacenCCN_ConsultasContablesAjax.asp?idDivision=" + division + "&idAlmacen=" + almacen + "&idCierre=" + idCierre + "&estadoCierre=" + estadoCierre + "&tipoCuenta=" + tipoCuenta + "&fecha=" + fecha + "&cdCuenta=" + cdCuenta + "&idSector=" + idSector + "&idObra=" + idObra + "&idBudgetArea=" + idBudgetArea + "&idBudgetDetalle=" + idBudgetDetalle + "&idCategoria=" + idCategoria + "&idArticulo=" + idArticulo, "realizarConsultaCallback()");
		ch.send();			
	}
	function realizarConsultaCallback(){
		document.getElementById("imgLoading").style.position = "absolute";
		document.getElementById("imgLoading").style.visibility  = "hidden";
		document.getElementById("lblLoading").style.position = "absolute";
		document.getElementById("lblLoading").style.visibility  = "hidden";
		document.getElementById("results").innerHTML = ch.response(); 
		var idCierre = 0;
		if (document.getElementById("idCierreAFirmar")) 
			idCierre = document.getElementById("idCierreAFirmar").value;
		link = "almacenCCN_FirmarAsientos.asp?idCierre=" + idCierre + "&secuencia=";
		hkey1 = new Hkey('hk1', link + "<%=FIRMA_ROL_RESP_CONTADURIA%>", '<% =HKEY() %>', 'hkey_callback()');
		hkey2 = new Hkey('hk2', link + "<%=FIRMA_ROL_RESP_PUERTO%>", '<% =HKEY() %>', 'hkey_callback()');
		hkey1.start();
		hkey2.start();	
	}	
	function hkey_callback(resp){
		if (resp != "<% =RESPUESTA_OK %>") {
			alert(resp);
		}
		else{
			submitPage();
		}
	}
	function addFechaCierre(pFecha, pCierre, pEstadoCierre, pTipoCuenta){
		document.getElementById("txtMesCierre").value = pFecha.substring(4,6);
		document.getElementById("txtAnioCierre").value = pFecha.substring(0,4);
		//document.getElementById("fechaCierre").value = pFecha;
		document.getElementById("idCierre").value = pCierre;
		document.getElementById("estadoCierre").value = pEstadoCierre;
		document.getElementById("tipoCuenta").value = pTipoCuenta;
		submitPage();
	}
	function delFecha(){
		document.getElementById("txtMesCierre").value = "";
		document.getElementById("txtAnioCierre").value = "";
		//document.getElementById("fechaCierre").value = "";
		document.getElementById("idCierre").value = "";
		document.getElementById("estadoCierre").value = "";
		document.getElementById("tipoCuenta").value = "";
		document.getElementById("cdCuenta").value = "";
		document.getElementById("cdObra").value = "";	
		document.getElementById("idObra").value = "";	
		document.getElementById("idSector").value = "";
		document.getElementById("idBudgetArea").value = "";
		document.getElementById("idBudgetDetalle").value = "";
		document.getElementById("cdCategoria").value = "";
		document.getElementById("idCategoria").value = "";
		document.getElementById("idArticulo").value = "";
		submitPage();
	}
	function addCuenta(pCuenta){
		document.getElementById("cdCuenta").value = pCuenta;
		submitPage();
	}
	function delCuenta(){
		document.getElementById("cdCuenta").value = "";
		document.getElementById("cdObra").value = "";
		document.getElementById("idObra").value = "";
		document.getElementById("idSector").value = "";
		document.getElementById("idBudgetArea").value = "";
		document.getElementById("idBudgetDetalle").value = "";
		document.getElementById("cdCategoria").value = "";
		document.getElementById("idCategoria").value = "";
		document.getElementById("idArticulo").value = "";
		submitPage();
	}
	function addObra(pIdObra, pCdObra, pIdBudgetArea, pIdBudgetDetalle){
		document.getElementById("cdObra").value = pCdObra;
		document.getElementById("idObra").value = pIdObra;
		document.getElementById("idBudgetArea").value = pIdBudgetArea;
		document.getElementById("idBudgetDetalle").value = pIdBudgetDetalle;
		submitPage();
	}
	function delObra(){
		document.getElementById("cdObra").value = "";
		document.getElementById("idObra").value = "";
		document.getElementById("idSector").value = "";
		document.getElementById("idBudgetArea").value = "";
		document.getElementById("idBudgetDetalle").value = "";
		document.getElementById("cdCategoria").value = "";
		document.getElementById("idCategoria").value = "";
		document.getElementById("idArticulo").value = "";
		submitPage();
	}		
	function addSector(pIdSector){
		document.getElementById("idSector").value = pIdSector;
		submitPage();
	}
	function delSector(){
		document.getElementById("idSector").value = "";
		document.getElementById("cdCategoria").value = "";
		document.getElementById("idCategoria").value = "";
		document.getElementById("idArticulo").value = "";
		submitPage();
	}		
	function addCategoria(pIdCategoria, pCdCategoria){
		document.getElementById("cdCategoria").value = pCdCategoria;
		document.getElementById("idCategoria").value = pIdCategoria;
		submitPage();
	}
	function delCategoria(){
		document.getElementById("cdCategoria").value = "";
		document.getElementById("idCategoria").value = "";
		document.getElementById("idArticulo").value = "";
		submitPage();
	}	
	function addArticulo(pIdArticulo){
		document.getElementById("idArticulo").value = pIdArticulo;
		submitPage();
	}
	function delArticulo(){
		document.getElementById("idArticulo").value = "";
		submitPage();
	}
	function openVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function editVale(id) {
		window.open("almacenValesTitulo.asp?idVale=" + id, "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>
<form id="frmSel">
<table class="reg_Header2" align="center" width="90%"  border="0">
		<%
		if flagUno then
		%>
				<tr><td colspan="2">&nbsp</td></tr>
		<%
		else
		%>
			<tr>
				<td>
					<font class="big2"><%=GF_TRADUCIR("División:")%></font>					
					
						<%
						strSQL = "SELECT * FROM TBLDIVISIONES WHERE"
						listDivi = getListaCargosAdmin()
						if (listDivi <> "") then
							 strSQL = strSQL & " IDDIVISION IN (" & listDivi & ")"
						else
							strSQL = strSQL & " IDDIVISION = -1"	'No tiene permisos para ninguna division
						end if
						Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
						if idDivision = 0 then mySelected = "SELECTED"
						%>
						<select onchange="delFecha();" name="idDivision" id="idDivision">
						<option title="Seleccione..." VALUE="0" <%=mySelected%>>Seleccione...</option><%
						while not rs.eof
							mySelected = ""
							if rs("IDDIVISION") = idDivision then mySelected = "SELECTED"
							%><option title="<%=rs("DSDIVISION")%>" VALUE="<%=rs("IDDIVISION")%>" <%=mySelected%>><%=rs("DSDIVISION")%></option><%
							rs.movenext
						wend
						mySelected = ""
						if idDivision = 1 then mySelected = "SELECTED"
						%>
						<option title="Exportacion" VALUE="1" <%=mySelected%>><%=GF_Traducir("EXPORTACION (EXPO)")%></option>
						<%
						Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
						%>
					</select>		
				</td>
				<% 
				if tipoCierre = TIPO_CIERRE_DEFINITIVO then 
					Response.Write "<td align='right'><font id='marquee' color='red' class='BIG'>&nbsp;ATENCION! CIERRE DEFINITIVO&nbsp;</font></td>"
				end if
				%>
			</tr>	
		<%
		end if
		%>	
		<input type="hidden" name="idAlmacen2" id="idAlmacen2" <%=pIdAlmacen2%>>	
</table>		
<input type="hidden" id="txtMesCierre"		name="txtMesCierre"		value="<%=txtMesCierre%>">
<input type="hidden" id="txtAnioCierre"		name="txtAnioCierre"	value="<%=txtAnioCierre %>">
<input type="hidden" id="fechaCierre"		name="fechaCierre"		value="<%=fechaCierre%>">
<input type="hidden" id="idCierre"			name="idCierre"			value="<%=idCierre%>">
<input type="hidden" id="estadoCierre"		name="estadoCierre"		value="<%=estadoCierre%>">
<input type="hidden" id="tipoCuenta"		name="tipoCuenta"		value="<%=tipoCuenta%>">
<input type="hidden" id="cdCuenta"			name="cdCuenta"			value="<%=cdCuenta%>">
<input type="hidden" id="cdObra"			name="cdObra"			value="<%=cdObra%>">
<input type="hidden" id="idObra"			name="idObra"			value="<%=idObra%>">
<input type="hidden" id="idSector"			name="idSector"			value="<%=idSector%>">
<input type="hidden" id="idBudgetArea"		name="idBudgetArea"		value="<%=idBudgetArea%>">
<input type="hidden" id="idBudgetDetalle"	name="idBudgetDetalle"	value="<%=idBudgetDetalle%>">
<input type="hidden" id="cdCategoria"		name="cdCategoria"		value="<%=cdCategoria%>">
<input type="hidden" id="idCategoria"		name="idCategoria"		value="<%=idCategoria%>">
<input type="hidden" id="idArticulo"		name="idArticulo"		value="<%=idArticulo%>">
<table align="center" align="center" width="90%" border="0">
	<tr>
		<td align="right">
			<%	
			if idArticulo <> 0 then
				if flagLink then
					strLink = "<img src='images/arrow_categ.gif'><a class='link' onclick='delArticulo()' title='" & dsArticulo & "'>" & idArticulo & "</a>" & strLink
				else
					strLink = "<img src='images/arrow_categ.gif'><font title='" & dsArticulo & "'>" & idArticulo & "</font>"
					flagLink = true
				end if
			end if
			if cdCategoria <> "" then
				if flagLink then
					strLink = "<img src='images/arrow_categ.gif'><a class='link' onclick='delArticulo()' title='" & dsCategoria & "'>" & cdCategoria & "</a>" & strLink
				else
					strLink = "<img src='images/arrow_categ.gif'><font title='" & dsCategoria & "'>" & cdCategoria & "</font>"
					flagLink = true
				end if
			end if
			if idSector <> 0 then
				if flagLink then
					strLink = "<img src='images/arrow_categ.gif'><a class='link' onclick='delCategoria()' title='" & dsSector & "'>" & idSector & "</a>" & strLink
				else
					strLink = "<img src='images/arrow_categ.gif'><font title='" & dsSector & "'>" & idSector & "</font>"
					flagLink = true
				end if
			end if
			if cdObra <> "" then
				if flagLink then
					strLink = "<img src='images/arrow_categ.gif'><a class='link' onclick='delSector()' title='Obra:" & cdObra & "&nbsp;-&nbsp;Area:" & idBudgetArea & "&nbsp; Detalle:" & idBudgetDetalle & "'>" & cdObra & "&nbsp;" & idBudgetArea & "-" & idBudgetDetalle & "</a>" & strLink
				else
					strLink = "<img src='images/arrow_categ.gif'><font title='Obra:" & cdObra & "&nbsp;-&nbsp;Area:" & idBudgetArea & "&nbsp; Detalle:" & idBudgetDetalle & "'>" & cdObra & "&nbsp;" & idBudgetArea & "-" & idBudgetDetalle & "</font>"
					flagLink = true
				end if
			end if
			if cdCuenta <> "" then
				if flagLink then
					strLink = "<img src='images/arrow_categ.gif'><a class='link' onclick='delObra()' title='" & cdCuenta & "'>" & formatCuentaPantalla(cdCuenta) & "</a>" & strLink
				else
					strLink = "<img src='images/arrow_categ.gif'><font title='" & cdCuenta & "'>" & formatCuentaPantalla(cdCuenta) & "</font>" 
					flagLink = true
				end if
			end if
			if txtFechaCierre <> "" then
				if flagLink then
					strLink = "<a class='link' onclick='delCuenta()' title='" & txtAnioCierre & "-" & txtMesCierre & "'>" & txtAnioCierre & "-" & txtMesCierre & "</a>" & strLink
				else
					strLink = "<font title='" & txtMesCierre & "-" & txtAnioCierre & "'>" & txtMesCierre & "-" & txtAnioCierre & "</font>"
					flagLink = true
				end if
			end if	
			if flagLink then
				strLink	=	"<a class='link' onclick='delFecha()' title='" & GF_TRADUCIR("Seleccionar otra fecha de cierre") & "'>CIERRES</a><img src='images/arrow_categ.gif'>" & strLink
			end if	
			Response.Write strLink		
			%>
		</td>
	</tr>
</table>
<table align="center" align="center" width="90%" border="0">
	<tr>
		<td align="center">
			<img style="position:absolute;visibility:hidden;" id="imgLoading" src="images/Loading4.gif">
			<div id="lblLoading"><b><br>Aguarde por favor...</b></div>
			<div id="results"></div>
		</td>
	</tr>
</table>
</form>
</body>
</html>