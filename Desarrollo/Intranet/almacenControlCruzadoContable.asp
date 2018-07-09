<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<% 
Call initAccessInfo(RES_ACC_AL)

dim idAlmacen, rs, idArticulo, dsArticulo, cdCategoria, dsCategoria
dim txtMesCierre, txtAnioCierre, txtFechaCierre, idDivision
idDivision = GF_Parametros7("idDivision", 0, 6)
txtMesCierre = GF_Parametros7("txtMesCierre", "", 6)
txtAnioCierre = GF_Parametros7("txtAnioCierre", "", 6)
idCierre = GF_PARAMETROS7("idCierre",0,6)
txtFechaCierre = txtAnioCierre & txtMesCierre
idArticulo = GF_PARAMETROS7("idArticulo",0,6)
cdCategoria = GF_PARAMETROS7("cdCategoria","",6)
idCategoria = GF_PARAMETROS7("idCategoria",0,6)

if idArticulo <> 0 then call getArticuloFull(idArticulo, dsArticulo, "")
if idCategoria <> 0 then dsCategoria = getCategoriaDS(idCategoria)
idAlmacen = getAlmacenesPorDivision(idDivision)
%>
<html>
<head>
<title>Almacenes</title>
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
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var ch = new channel();	
	var link;
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
		tb.addButton("Contabilidad_Folder-16x16.png", "Contabilidad", "irA('almacenContabilidad.asp')");				
		tb.draw();		
		realizarConsulta('<%=idDivision%>', '<%=idAlmacen%>', '<%=txtFechaCierre%>');
		pngfix();

	}
	function submitPage(){
		document.getElementById("frmSel").submit();
	}
	
	function realizarConsulta(division, almacen, fecha){
		document.getElementById("imgLoading").style.position = "relative";
		document.getElementById("imgLoading").style.visibility  = "visible";
		ch.bind("almacenCC_ControlCruzadoContableAjax.asp?idDivision=" + division + "&idAlmacen=" + almacen + "&fecha=" + fecha, "realizarConsultaCallback()");
		ch.send();			
	}
	function realizarConsultaCallback(){
		document.getElementById("imgLoading").style.position = "absolute";
		document.getElementById("imgLoading").style.visibility  = "hidden";
		document.getElementById("results").innerHTML = ch.response(); 
	}	
	function verCruzado(pFecha, pIdCierre){
		document.getElementById("imgLoading").style.position = "relative";
		document.getElementById("imgLoading").style.visibility  = "visible";
		document.getElementById("results").innerHTML = ""; 
		ch.bind("almacenCC_ControlCruzadoContableAjax.asp?idCierre=" + pIdCierre + "&idAlmacen=<%=idAlmacen%>&idDivision=<%=idDivision%>&fecha=" + pFecha, "realizarConsultaCallback()");
		ch.send();			
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Administración Almacenes - Control Cruzado Contable (CCC)") %>
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
					<select onchange="submitPage();" name="idDivision" id="idDivision">
						<%
						strSQL = "SELECT * FROM TBLDIVISIONES WHERE IDDIVISION IN (" & getListaCargosAdmin & ")"
						call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
						if idDivision = 0 then mySelected = "SELECTED"
						%><option title="Seleccione..." VALUE="0" <%=mySelected%>>Seleccione...</option><%
						while not rs.eof
							mySelected = ""
							if rs("IDDIVISION") = idDivision then mySelected = "SELECTED"
							%><option title="<%=rs("DSDIVISION")%>" VALUE="<%=rs("IDDIVISION")%>" <%=mySelected%>><%=rs("DSDIVISION")%></option><%
							rs.movenext
						wend
						call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
						%>
					</select>		
				</td>
			</tr>	
		<%
		end if
		%>	
</table>		
<input type="hidden" id="txtMesCierre"		name="txtMesCierre"		value="<%=txtMesCierre%>">
<input type="hidden" id="txtAnioCierre"		name="txtAnioCierre"	value="<%=txtAnioCierre %>">

<table align="center" align="center" width="90%" border="0">
	<tr>
		<td align="center">
			<img style="position:absolute;visibility:hidden;" id="imgLoading" src="images/Loading4.gif">
			<div id="results"></div>
		</td>
	</tr>
</table>
</form>
</body>
</html>