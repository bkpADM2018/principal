<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
Call initAccessInfo(RES_ACC_AL)

Dim pidDivision, rs, titleAux, idArticulo, dsArticulo, cdCategoria, cdVale, cierresList, changeDiv
dim ultimoCierre, ultimoCierreFN, proximoCierre, proximoCierrePRO, idAlmacenes, tipoCierre, flagPuedeCerrar, totalDivision
DIM anio, mes, dia, anioUltimoDef, mesUltimoDef, idCierre, cotizacionDolar, reasonCode
dim almacenesArroyo, almacenesTransito, almacenesBahia
dim fecAux, mesAux
'Hardcodeo para las siglas de los puertos que se muestran en pantalla
dim siglasArroyo, siglasTransito, siglasBahia, siglasExpo
siglasArroyo = "ARR"
siglasTransito = "TRA"
siglasBahia = "BBA"
siglasExpo = "EXP"
'Hardcodeo para los Id de divisiones
dim ARROYO, TRANSITO, BAHIA, EXPO
ARROYO = 2
TRANSITO = 4
BAHIA = 3
EXPO = 1

'Obtener valores submitidos
tipoCierre = GF_Parametros7("tipoCierre", "", 6)
cotizacionDolar = GF_Parametros7("cotizacionDolar", 3, 6)
proximoCierrePRO = GF_Parametros7("proximoCierrePRO","",6)
'proximoCierrePRO = "20130131"
if tipoCierre = "" then tipoCierre = TIPO_CIERRE_PROVISORIO

'Cargar lista de almacenes para cada division
strSQL = "SELECT * FROM TBLALMACENES WHERE ESTADO=" & ESTADO_ACTIVO
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof 
	select case CINT(rs("IDDIVISION"))
		case ARROYO 
			almacenesArroyo = rs("IDALMACEN") & "," & almacenesArroyo
		case TRANSITO
			almacenesTransito = rs("IDALMACEN") & "," & almacenesTransito    
		case BAHIA
			almacenesBahia = rs("IDALMACEN") & "," & almacenesBahia    
	end select		
	rs.movenext
wend
if len(almacenesArroyo)>0 then almacenesArroyo = left(almacenesArroyo,len(almacenesArroyo)-1)
if len(almacenesTransito)>0 then almacenesTransito = left(almacenesTransito,len(almacenesTransito)-1)
if len(almacenesBahia)>0 then almacenesBahia = left(almacenesBahia,len(almacenesBahia)-1)

'Obtener proxima fecha de cierre sugerida
if tipoCierre = TIPO_CIERRE_PROVISORIO then
	'Si quiere hacer un cierre Provisorio se busca el ultimo Definitivo para obtener cual es el siguiente a realizar.
	flagPuedeCerrar = true
	strSQL = "SELECT MAX(ANIO * 100 + MES) AS ULTIMOCIERRE FROM TBLCIERRESCABECERA2 WHERE IDDIVISION=" & ARROYO & " AND ESTADO='" & TIPO_CIERRE_DEFINITIVO & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not IsNull(rs("ULTIMOCIERRE")) then
		anio = left(rs("ULTIMOCIERRE"),4)
		mes = mid(rs("ULTIMOCIERRE"),5,2)
	else
		anio = "2011"
		mes = "01"
	end if
	anioUltimoDef = anio
	mesUltimoDef = mes
	'if proximoCierrePRO = "" then 
		'No se eligio ninguna fecha aun, sugerir la encontrada en las lineas anteriores.
		if len(mes) = 1 then mes = "0" & mes
		dia = getLastDayOfMonth(anio,mes)	
		'Cargar la fecha del ultimo cierre Definitivo encontrado
		ultimoCierre = GF_FN2DTE(anio & mes & dia)
		ultimoCierreFN = anio & mes & dia			
		'Cargar la fecha del proximo cierre Provisorio
		if mes=12 then 
			anio = anio + 1
			mes = 1
		else	
			mes = mes + 1 
		end if	
		if len(mes) = 1 then mes = "0" & mes
		proximoCierrePRO = anio & mes & getLastDayOfMonth(anio,mes)
		proximoCierreFN = proximoCierrePRO		
		proximoCierre = GF_FN2DTE(proximoCierrePRO)		
else
	'Si quiere hacer un cierre Definitivo se busca el ultimo Definitivo y solo se podra cerrar definitivamente el del mes siguiente a ese.
	strSQL = "SELECT MAX(ANIO * 100 + MES) AS ULTIMOCIERRE FROM TBLCIERRESCABECERA2 WHERE IDDIVISION=" & ARROYO & " AND ESTADO='" & TIPO_CIERRE_DEFINITIVO & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not IsNull(rs("ULTIMOCIERRE")) then
		anio = left(rs("ULTIMOCIERRE"),4)
		mes = mid(rs("ULTIMOCIERRE"),5,2)
	else
		anio = "2010"
		mes = "01"
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	anioUltimoDef = anio
	mesUltimoDef = mes
	dia = getLastDayOfMonth(anio,mes)
	if len(mes) = 1 then mes = "0" & mes
	ultimoCierre = GF_FN2DTE(anio & mes & dia)
	ultimoCierreFN = anio & mes & dia	
	'Se carga el mes sigueinte al ultimo cierre defitivo realizado.
	proximoCierre = DateAdd("m", 1, ultimoCierre)
	anio = year(proximoCierre)
	mes = month(proximoCierre)
	
	flagPuedeCerrar = estaFirmado(ARROYO,anio,mes)
	if flagPuedeCerrar then 
		flagPuedeCerrar = estaFirmado(TRANSITO,anio,mes)
		if flagPuedeCerrar then 
			flagPuedeCerrar = estaFirmado(BAHIA,anio,mes)
		end if
	end if	
	if not (flagPuedeCerrar) then reasonCode = flagPuedeCerrar

	'Se arma nuevamente la fecha del cierre provisorio y se calcula el total para mostrar en pantalla. Solo informativo, Revisar el tema de las fechas!
	if len(mes) = 1 then mes = "0" & mes
	dia = getLastDayOfMonth(anio,mes)
	proximoCierre = GF_FN2DTE(anio & mes & dia)
	proximoCierreFN = anio & mes & dia
	idCierreArr = getIdCierre2(anio, mes, ARROYO, TIPO_CIERRE_PROVISORIO)
	totalDivisionArr = getTotalPorCierre(idCierreArr)
	idCierreTra = getIdCierre2(anio, mes, TRANSITO, TIPO_CIERRE_PROVISORIO)
	totalDivisionTra = getTotalPorCierre(idCierreTra)
	idCierreBba = getIdCierre2(anio, mes, BAHIA, TIPO_CIERRE_PROVISORIO)
	totalDivisionBba = getTotalPorCierre(idCierreBba)
	idCierreExp = getIdCierre2(anio, mes, EXPO, TIPO_CIERRE_PROVISORIO)
	totalDivisionExp = getTotalPorCierre(idCierreExp)
	'if cdbl(totalDivisionExp) = 0 then call actualizarEstadoCierre(idCierreExp, TIPO_CIERRE_DEFINITIVO)
end if
'----------------------------------------------------------------------------------------
function getTotalPorCierre(pIdCierre)
dim strSQL, rs, oConn, rtrn
rtrn = 0
	strSQL ="Select SUM(IMPORTEPESOS) TOTALPESOS, SUM(IMPORTEDOLARES) TOTALDOLARES " &_
	        " from (Select IDCIERRE, CDCUENTA, CCOSTOS, DBCR, round(SUM(IMPORTEPESOS)/10000, 2) IMPORTEPESOS, round(SUM(IMPORTEDOLARES)/10000, 2) IMPORTEDOLARES from TBLCIERRESASIENTOS2 group by IDCIERRE, CDCUENTA, CCOSTOS, DBCR) T" & _ 
			"    where idcierre=" & pIdCierre & " and dbcr= " & TIPO_CIERRE_DEBE
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not isNull(rs("TOTALPESOS")) then rtrn = cdbl(rs("TOTALPESOS"))*100
getTotalPorCierre = rtrn
end function
'----------------------------------------------------------------------------------------
function estaFirmado(pDivision, pAnio, pMes)
dim strSQL, rs, oConn, rtrn
rtrn = false
	strSQL = "SELECT * FROM TBLCIERRESCABECERA2 WHERE IDDIVISION=" & pDivision & " AND ESTADO='" & TIPO_CIERRE_PROVISORIO & "' AND ANIO=" & pAnio & " AND MES=" & pMes
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		strSQL = "SELECT * FROM TBLCIERRESFIRMAS2 WHERE IDCIERRE=" & rs("IDCIERRE") & " AND HKEY IS NULL"
		'Response.Write STRSQL
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if rs.eof then 
			rtrn = true
		else
			rtrn = 2	
		end if	
	else
		rtrn = 1
	end if
estaFirmado = rtrn
end function
'----------------------------------------------------------------------------------------
%>
<html>
<head>
<title>Almacenes - Cierres Contables</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
    <link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<style type="text/css">
.link {
	cursor:pointer;
	color:blue;
	text-decoration:underline;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>	
<script type="text/javascript" src="scripts/script_fechas.js"></script>
    <script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
    <script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
	var ch = new channel();	
	var numberOfFase = -1;
	var flagLogSF = "";
	var flagErrorSF = false;
	function check_callback(resp) {				
		//if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}
	
	function irA(pLink) {
		location.href = pLink;
	}
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}	
	function marquee(){
		if (document.getElementById("marquee").style.color == 'red'){
			document.getElementById("marquee").style.color = 'white'; 
		}else{
			document.getElementById("marquee").style.color = 'red'; 
		}	
	}
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.addButton("refresh-16x16.png", "Refresh", "submitPage()");
		tb.addButton("Contabilidad_Folder-16x16.png", "Contabilidad", "irA('almacenCCN_Contabilidad.asp')");		
		tb.draw();		
		<%if tipoCierre = TIPO_CIERRE_DEFINITIVO and reasonCode = 0 then%>	
			setInterval("marquee()",500);
			resaltarFase(0,2);
			resaltarSubFase(0+'a',2);
			resaltarSubFase(0+'t',2);
			resaltarSubFase(0+'b',2);
			resaltarFase(1,2);
			resaltarSubFase(1+'a',2);
			resaltarSubFase(1+'t',2);
			resaltarSubFase(1+'b',2);			
			resaltarFase(2,2);
			resaltarSubFase(2+'a',2);
			resaltarSubFase(2+'t',2);
			resaltarSubFase(2+'b',2);			
			resaltarFase(3,2);
			resaltarSubFase(3+'f',2);
			resaltarSubFase(3+'fa',2);
			resaltarSubFase(3+'ft',2);
			resaltarSubFase(3+'fb',2);
			resaltarSubFase(3+'r',2);
			resaltarSubFase(3+'ra',2);
			resaltarSubFase(3+'rt',2);
			resaltarSubFase(3+'rb',2);									
			resaltarSubFase(3+'t',2);
			resaltarSubFase(3+'tt',2);
			resaltarFase(4,2);
			resaltarSubFase(4+'a',2);
			resaltarSubFase(4+'t',2);
			resaltarSubFase(4+'b',2);			
			resaltarFase(5,2);
			resaltarSubFase(5+'a',2);
			resaltarSubFase(5+'t',2);
			resaltarSubFase(5+'b',2);				
			resaltarFase(6,2);
			resaltarSubFase(6+'a',2);
			resaltarSubFase(6+'t',2);
			resaltarSubFase(6+'b',2);	
			
			//agregarATasks("Inicializando...");
			agregarATasks("Gastos realizados en el mes por la division Arroyo: $ " + document.getElementById("totalDivisionArr").value);
			agregarATasks("Gastos realizados en el mes por la division Transito: $ " + document.getElementById("totalDivisionTra").value);
			agregarATasks("Gastos realizados en el mes por la division Bahia: $ " + document.getElementById("totalDivisionBba").value);
			agregarATasks("Gastos realizados en el mes por la division Exportacion: $ " + document.getElementById("totalDivisionExp").value);
			//agregarATasks("Inicializacion exitosa...");
		<%end if %>			
		<%if tipoCierre = TIPO_CIERRE_DEFINITIVO then%>	
			SeleccionarCalLimite(undefined, '<% = proximoCierre %>');
		<% end if %>
	}
	function submitPage(changeDiv){
		if (changeDiv==1) document.getElementById("changeDiv").value = '1'; 
		if (document.getElementById("cotizacionDolar")) {
			document.getElementById("cotizacionDolar").value = ""; 
		}
		document.getElementById("frmSel").submit();
	}
	function openVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	
	/*PRE-INICIALIZACION*/
	function preInicializacion(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase,0);
		resaltarSubFase(numberOfFase + 'a',0);
		resaltarSubFase(numberOfFase + 't',0);
		resaltarSubFase(numberOfFase + 'b',0);
		agregarATasks("PreInicializando...");
		ch.bind("almacenCCN_PreInicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "preInicializacion_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_PreInicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "preInicializacion_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_PreInicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "preInicializacion_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();
	}
	var flagErrorPreIni = false;
	function preInicializacion_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		var str = new String();
		str = ch.response();
		//alert(str);
		descomposicion = str.split("-")
		if (descomposicion[0] != "V")flagErrorPreIni = true
		if (pDivision == '<%=siglasBahia%>'){		    
			if (flagErrorPreIni) {
				agregarATasks("PreInicializacion finlizada con errores...");
				if (confirm("El proceso de Pre-Inicializacion ha finalizado, desea ver el log generado?")){
	    		    var puw = new winPopUp('popUp',descomposicion[1], '800','400','Log Pre-Inicializacion');
			    }
			} else {
		    	agregarATasks("PreInicializacion finlizada con exito...");
			    resaltarFase(numberOfFase, 1);			    			    
			    inicializacion();				    
			}
		}
	}
	
	/*INICIALIZACION*/
	function inicializacion(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase,0);
		resaltarSubFase(numberOfFase + 'a',0);
		resaltarSubFase(numberOfFase + 't',0);
		resaltarSubFase(numberOfFase + 'b',0);		
		agregarATasks("--------------------------------------------------------------------------");
		agregarATasks("Inicializando...");
		ch.bind("almacenCCN_InicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "inicializacion_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_InicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "inicializacion_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_InicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=EXPO%>&idAlmacen=0&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "inicializacion_Callback('" + numberOfFase + "t','<%=siglasExpo%>')");
		ch.send();
		ch.bind("almacenCCN_InicializacionAjax.asp?tipoCierre=<%=tipoCierre%>&idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "inicializacion_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();
	}
	function inicializacion_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Inicializacion para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
				//valuacioContableFacturas();
				stockFisico();
		}		
	}		

	/*STOCK FISICO*/
	function stockFisico(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase, 0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Iniciando cierre de stock fisico...");
		callAjax('almacenCCN_StockFisicoAjax.asp?fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>', '<% =almacenesArroyo %>', 'stockFisico', numberOfFase + 'a', '<%=siglasArroyo%>')
		callAjax('almacenCCN_StockFisicoAjax.asp?fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>', '<% =almacenesTransito %>', 'stockFisico', numberOfFase + 't', '<%=siglasTransito%>')
		callAjax('almacenCCN_StockFisicoAjax.asp?fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>', '<% =almacenesBahia %>', 'stockFisico', numberOfFase + 'b', '<%=siglasBahia%>')
	}
	function stockFisico_Callback(pAlm, pSubFase, pDivision){
		if (flagErrorSF){
				agregarATasks("Cierre de Stock Fisico falló para el almacen " + pAlm + "...");
				if (confirm("El cierre del Stock Fisico ha fallado para el almacen " + pAlm + ". Desea ver el log?") == true){
					popUp = new PopUpWindow('popUp', flagLogSF, '800', '390', 'Logs');
				}	
		} 
		else{
			resaltarSubFase(pSubFase, 1);
			agregarATasks("Almacen " + pAlm + "...OK!");	
			if (pDivision == '<%=siglasBahia%>'){
					resaltarFase(numberOfFase, 1);
					valuacioContableFacturas();
					//gastos();
			}	
		}
	}	

	function callAjax(link, idAlmacenes, funcCB, pSubFase, pDivision) {				
		link += '&idAlmacenes=';
		var arrAlmacenes = idAlmacenes.split(",");				
		//agregarATasks("Total " + pDivision + "---" + arrAlmacenes.length);		
		while (arrAlmacenes.length > 0) {	
			alm = arrAlmacenes.pop();
			//alert(link + alm);
			if (arrAlmacenes.length == 0) {
				ch.bind(link + alm, funcCB + "_Callback('" + alm + "','" + pSubFase + "','" + pDivision + "')");
				agregarATasks("Stock Fisico para " + alm + " exitoso...");			
			} else {
				ch.bind(link + alm, "callAjax_Callback('" + alm + "')");
						
			}
			ch.send();
		}
	}
	function callAjax_Callback(pAlm) {
		var str;
		var descomposicion;
		str = ch.response();
		//alert(str);
		if (str != "Hecho..."){
			flagErrorSF = true;
			descomposicion = str.split("-");
			flagLogSF = descomposicion[1];
		}		
		else{
			//agregarATasks(ch.response());										
			agregarATasks("Almacen " + pAlm + "...OK!");		
		}
	}

	
	/*VALUACION CONTABLE - FACTURAS*/
	function valuacioContableFacturas(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase,0);
		resaltarSubFase(numberOfFase + 'f',0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Valuando contablemente a partir de Facturas...");
		ch.bind("almacenCCN_ValuacionContableFacturasAjax.asp?idDivision=<%=ARROYO%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableFacturas_Callback('" + numberOfFase + "fa','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_ValuacionContableFacturasAjax.asp?idDivision=<%=TRANSITO%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableFacturas_Callback('" + numberOfFase + "ft','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_ValuacionContableFacturasAjax.asp?idDivision=<%=BAHIA%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableFacturas_Callback('" + numberOfFase + "fb','<%=siglasBahia%>')");
		ch.send();
	}
	function valuacioContableFacturas_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Valuación Contable a partir de Facturas para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				//resaltarFase(numberOfFase, 1);
				resaltarSubFase(numberOfFase + 'f',1);
				valuacioContableVRS();
		}		
	}	
	
	/*VALUACION CONTABLE - VRS*/
	function valuacioContableVRS(){
		resaltarSubFase(numberOfFase + 'r',0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Valuando contablemente a partir de Reclasificaciones...");
		ch.bind("almacenCCN_ValuacionContableVRSAjax.asp?idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableVRS_Callback('" + numberOfFase + "ra','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_ValuacionContableVRSAjax.asp?idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableVRS_Callback('" + numberOfFase + "rt','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_ValuacionContableVRSAjax.asp?idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>", "valuacioContableVRS_Callback('" + numberOfFase + "rb','<%=siglasBahia%>')");
		ch.send();
	}
	function valuacioContableVRS_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Valuación Contable a partir de Reclasificaciones para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarSubFase(numberOfFase + 'r',1);
				valuacioContableVMR();
		}		
	}	
	/*VALUACION CONTABLE - VMR*/
	function valuacioContableVMR(){
		resaltarSubFase(numberOfFase + 't',0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Valuando contablemente a partir de Transferencias...");
		ch.bind("almacenCCN_ValuacionContableVMRAjax.asp?fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "valuacioContableVMR_Callback('" + numberOfFase + "t','<%=siglasArroyo%>')");
		ch.send();
	}
	function valuacioContableVMR_Callback(pSubFase, pDivision){
		resaltarFase(numberOfFase, 1);
		resaltarSubFase(pSubFase, 1);
		resaltarSubFase(pSubFase + 't', 1);
		agregarATasks("Valuacion Contable a partir de Transferencias para todas las divisiones exitosa...");
		//return 0;
		aplicacionDePrecios();
	}	
		
	/*APLICACION DE PRECIOS*/
	function aplicacionDePrecios(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase, 0);
		resaltarSubFase(numberOfFase,0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Aplicando Precios Contables a Vales...");
		ch.bind("almacenCCN_AplicacionDePreciosAjax.asp?idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "aplicacionDePrecios_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_AplicacionDePreciosAjax.asp?idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "aplicacionDePrecios_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_AplicacionDePreciosAjax.asp?idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "aplicacionDePrecios_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();
	}
	function aplicacionDePrecios_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Aplicacion de Precios para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
				armadoDeCtaCte();
		}	
	}			

	/*ARMADO DE CUENTA CORRIENTE*/
	function armadoDeCtaCte(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase, 0);
		resaltarSubFase(numberOfFase,0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Armando Cuenta Corriente Contable por Articulo...");
		ch.bind("almacenCCN_ArmadoCtaCteAjax.asp?idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "armadoDeCtaCte_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_ArmadoCtaCteAjax.asp?idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "armadoDeCtaCte_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_ArmadoCtaCteAjax.asp?idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "armadoDeCtaCte_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();

	}
	function armadoDeCtaCte_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Armado de Cuenta Corriente para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
				//valuacioContableFacturas();
				//stockFisico();
				gastos();
		}	
	}

			

	/*GASTOS*/	
	function gastos(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase,0);
		agregarATasks("--------------------------------------------------------------------------");		
		agregarATasks("Iniciando cierre gastos...");
		ch.bind("almacenCCN_AsientosAjax.asp?idDivision=<%=ARROYO%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "gastos_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosAjax.asp?idDivision=<%=TRANSITO%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "gastos_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosAjax.asp?idDivision=<%=EXPO%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "gastos_Callback('" + numberOfFase + "b','<%=siglasExpo%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosAjax.asp?idDivision=<%=BAHIA%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>", "gastos_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();

	}
	function gastos_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Cierre de Gastos para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
				finalizacion();
		}
	}
	
	/*FINALIZACION*/
	function finalizacion(){
		numberOfFase = numberOfFase + 1;
		resaltarFase(numberOfFase,0);
		agregarATasks("--------------------------------------------------------------------------");
		agregarATasks("Iniciando finalizacion...");
		ch.bind("almacenCCN_FinalizacionAjax.asp?idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>&tipoCierre=<%=tipoCierre%>&cotizacionDolar=" + document.getElementById("cotizacionDolar").value, "finalizacion_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_FinalizacionAjax.asp?idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>&tipoCierre=<%=tipoCierre%>&cotizacionDolar=" + document.getElementById("cotizacionDolar").value, "finalizacion_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_FinalizacionAjax.asp?idDivision=<%=EXPO%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>&tipoCierre=<%=tipoCierre%>&cotizacionDolar=" + document.getElementById("cotizacionDolar").value, "finalizacion_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_FinalizacionAjax.asp?idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&fechaCierre=<%=proximoCierreFN%>&fechaCierreAnt=<%=ultimoCierreFN%>&tipoCierre=<%=tipoCierre%>&cotizacionDolar=" + document.getElementById("cotizacionDolar").value, "finalizacion_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();
	}

	function finalizacion_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Finalizacion para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
		}		
	}

	/*CIERRE DEFINITIVO*/
	function inicializacionCierreDefinitivo(){
		numberOfFase = 7
		resaltarFase(numberOfFase,0);
		agregarATasks("Iniciando finalizacion...");
		//return 0;
		ch.bind("almacenCCN_AsientosContablesAjax.asp?idDivision=<%=ARROYO%>&idAlmacen=<%=almacenesArroyo%>&idCierre=<%=idCierreArr%>&fechaCierre=<%=proximoCierreFN%>&fechaAsiento=" + document.getElementById("closingdate").value, "inicializacionCierreDefinitivo_Callback('" + numberOfFase + "a','<%=siglasArroyo%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosContablesAjax.asp?idDivision=<%=TRANSITO%>&idAlmacen=<%=almacenesTransito%>&idCierre=<%=idCierreTra%>&fechaCierre=<%=proximoCierreFN%>&fechaAsiento=" + document.getElementById("closingdate").value, "inicializacionCierreDefinitivo_Callback('" + numberOfFase + "t','<%=siglasTransito%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosContablesAjax.asp?idDivision=<%=EXPO%>&idCierre=<%=idCierreExp%>&fechaCierre=<%=proximoCierreFN%>&fechaAsiento=" + document.getElementById("closingdate").value, "inicializacionCierreDefinitivo_Callback('" + numberOfFase + "b','<%=siglasExpo%>')");
		ch.send();
		ch.bind("almacenCCN_AsientosContablesAjax.asp?idDivision=<%=BAHIA%>&idAlmacen=<%=almacenesBahia%>&idCierre=<%=idCierreBba%>&fechaCierre=<%=proximoCierreFN%>&fechaAsiento=" + document.getElementById("closingdate").value, "inicializacionCierreDefinitivo_Callback('" + numberOfFase + "b','<%=siglasBahia%>')");
		ch.send();
	}

	function inicializacionCierreDefinitivo_Callback(pSubFase, pDivision){
		resaltarSubFase(pSubFase, 1);
		agregarATasks("Pasajes de Asientos para " + pDivision + " exitosa...");
		if (pDivision == '<%=siglasBahia%>'){
				resaltarFase(numberOfFase, 1);
		}
	}

	/*Inicio de cierres contables*/
	function faseStart(){
		<%if tipoCierre = TIPO_CIERRE_DEFINITIVO then%>	
			if (confirm("Esta seguro que desea realizar el cierre definitivo?")){
				inicializacionCierreDefinitivo();
			}
			else{
				return 0;
			}
		<%else%>	
			preInicializacion();
		<%end if%>		
		document.getElementById("cmdStart").disabled = true;
	}

	function resaltarFase(number, state){
		var color = "#dcdcdc";
		var colorFont = "#000000";
		var imgSrc = "images/loading2.gif";
		imgVisibility = "hidden";
		if (state == 1) {
			color = "#ffffcc";
			colorFont = "#000000";
			imgSrc = "images/icon_ok.gif";
			imgVisibility = "hidden";
		}	
		else if (state == 2) {
			color = "#dcdcdc";
			colorFont = "#778899";
			imgSrc = "images/1p.gif";
			imgVisibility = "hidden";
		}	
		else if (state == 3) {
			color = "#dcdcdc";
			colorFont = "#778899";
			imgSrc = "images/icon_del.gif";
			imgVisibility = "hidden";
		}			
		document.getElementById("fase" + number + "IMG").src = imgSrc;
		document.getElementById("fase" + number + "IMG").style.visibility = "visible";
		document.getElementById("fase" + number).style.backgroundColor = color;
		document.getElementById("fase" + number).style.color = colorFont;
	}

	function resaltarSubFase(number, state){
		var color = "#dcdcdc";
		var colorFont = "#000000";
		if (state == 1) {
			color = "#ffffcc";
			colorFont = "#000000";
		}	
		else if (state == 2) {
			color = "#dcdcdc";
			colorFont = "#778899";
		}	
		else if (state == 3) {
			color = "#ffffff";
			colorFont = "#778899";
		}	
		document.getElementById("fase" + number).style.backgroundColor = color;
		document.getElementById("fase" + number).style.color = colorFont;
	}

	function agregarATasks(dsTask){
		var objSelectTasks = document.getElementById("tasks");
		var optNueva = objSelectTasks.appendChild(document.createElement('option'));
		optNueva.text = dsTask;
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

	function CerrarCal(cal) {
		cal.hide();
	}

	function SeleccionarCalLimite(cal, date) {
		var str= new String(date);		
		document.getElementById("closingdateDiv").innerHTML = str;
	    document.getElementById("closingdate").value = str;
		if (cal) cal.hide();	
	}	

</script>
</head>
<body onLoad="bodyOnLoad()">
<% 
call GF_TITULO2("kogge64.gif","Administración Almacenes - Cierres Contables") %>
<div id="toolbar"></div>
<form id="frmSel">
<%
'if len(pidAlmacenes) = 0 then Response.End %>	
<br>
<% if tipoCierre = TIPO_CIERRE_DEFINITIVO then %>
<table class="reg_Header2" align="center" width="90%"  border="0">
	<tr>
		<td align='right'><font id='marquee' color='red' class='BIG'>&nbsp;ATENCION! CIERRE DEFINITIVO</font></td>
	</tr>	
</table>		
<% end if %>
<table align="center" width="90%" border="1" cellpadding=2 cellspacing=0>
	<tr class="reg_Header_nav">
		<td colspan="2"><%=GF_Traducir("Cierres Contables Unificados")%></td>
	</tr>
	<% if tipoCierre = TIPO_CIERRE_PROVISORIO then %>
	<tr>
		<td class="reg_Header_navdos" align="left" width="15%">
			<%=GF_Traducir("Fecha Cierre")%>
		</td>	
		<td>
		    <% =proximoCierre %>			
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos" align="left">
			<%=GF_Traducir("Cotizacion U$S")%>
		</td>	
		<td>
			<%
			mesAux = MONTH(proximoCierre)
			if len(mesAux) = 1 then mesAux = "0" & mesAux
			fecAux = GF_FN2DTCONTABLE(YEAR(proximoCierre) & mesAux & "01")
			if trim(cotizacionDolar) = 0 then
				strSQL = "SELECT AVG(TCAMBIO) AS AVG_COTIZACION FROM CGT012A WHERE FECHA BETWEEN '" & fecAux & "' AND '" & GF_FN2DTCONTABLE(proximoCierreFN) & "' AND CODMON='" & T_CAMBIO_COMPRADOR & "'"
				'Response.Write strSQL
				Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
				if not isNull(rs("AVG_COTIZACION")) then cotizacionDolar = GF_EDIT_DECIMALS(clng(cdbl(rs("AVG_COTIZACION"))*1000),3)
				Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "CLOSE", strSQL)
			end if
			%>
			<input style="text-align:right;" type="text" size="5" name="cotizacionDolar" id="cotizacionDolar" value="<%=cotizacionDolar%>">
		</td>
	</tr>			
	<% 
	else 
	%>
	<tr>
		<td class="reg_Header_navdos" align="left" width="18%">
			<%=GF_Traducir("Ultimo cierre")%>
		</td>	
		<td>
			&nbsp;<%=ultimoCierre%>
		</td>
	</tr>

	<tr>
		<td class="reg_Header_navdos" align="left">
			<%=GF_Traducir("Próximo")%>
		</td>	
		<td>
			&nbsp;<%=proximoCierre%>
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos" align="left">
			<%=GF_Traducir("Fecha de los Asientos")%>
		</td>	
		<td>
			<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/DATE.gif"></a>
			<span id="closingdateDiv"><% =proximoCierre %></span>						
			<input type="hidden" id="closingdate" name="closingdate" value="<% =proximoCierre %>" />
		</td>
	</tr>	
	<% end if %>

	<tr>
		<td class="reg_Header_navdos" align="left">
			<%=GF_Traducir("Tipo de Cierre")%>
		</td>	
		<td>
			<input onClick="submitPage()" style="cursor:pointer;" value="<%=TIPO_CIERRE_PROVISORIO%>" type="radio" id="tipoCierre" name="tipoCierre" <%if tipoCierre=TIPO_CIERRE_PROVISORIO then Response.Write "Checked"%>><%=GF_TRADUCIR("Provisorio")%>
			<input onClick="submitPage()" style="cursor:pointer;" value="<%=TIPO_CIERRE_DEFINITIVO%>" type="radio" id="tipoCierre" name="tipoCierre" <%if tipoCierre=TIPO_CIERRE_DEFINITIVO then Response.Write "Checked"%>><%=GF_TRADUCIR("Definitivo")%>
		</td>
	</tr>	
	<% 
	if tipoCierre = TIPO_CIERRE_DEFINITIVO then myStatus = "disabled"
	%>
	<tr>
		<td valign="top" class="reg_Header_navdos" align="left">
			<%=GF_Traducir("Resúmen")%>
		</td>	
		<td>
				<%
				if tipoCierre=TIPO_CIERRE_PROVISORIO then 
					Response.Write "Se realizará el cierre contable <b>PROVISORIO</b> al <b>" & proximoCierre & "</b> tomando como referencia anterior, el cierre contable realizado el <b>" & ultimoCierre & "</b> y todos los movimientos del mes de <b>" & getNameOfMonth(month(proximoCierre)) & "."
				else
					if not flagPuedeCerrar then
						select case(reasonCode)
							case 1 
								Response.Write "<font color='red'>Imposible realizar el cierre <b>DEFINITIVO</b> al <b>" & proximoCierre & "</b> debido a que no se cuenta con el cierre <b>PROVISORIO</b>. Por favor, realice el cierre <b>PROVISORIO</b> de dicha fecha e intentelo nuevamente.</font>"
							case 2
								Response.Write "<font color='red'>Imposible pasar a <b>DEFINITIVO</b> el cierre contable realizado el <b>" & proximoCierre & "</b> debido a que no se cuenta con las firmas de aprobacion pertinentes. Por favor, solicite la firma del cierre contable <b>PROVISORIO</b> de dicha fecha e intentelo nuevamente.</font>"						
							case 3
								Response.Write "<font color='red'>Imposible realizar el cierre <b>PROVISORIO</b> al <b>" & proximoCierre & "</b> debido a que dicho cierre ya se encuentra firmado por los responsables.</font>"
							case 4
								Response.Write "<font color='red'>Imposible realizar el cierre <b>DEFINITIVO</b> al <b>" & proximoCierre & "</b> para EXPORTACION. Para poder realizar dicho cierre es imprescindible que se encuentren firmados los cierres de las demas divisiones para este año y mes.</font>"
						end select	
					else
						Response.Write "Se pasará a <b>DEFINITIVO</b> el cierre contable realizado el <b>" & proximoCierre & "</b>."
					end if	
				end if
				%>
		</td>
	
	</tr>	
	<tr>
		<td colspan="2" align="center">
			<input type="button" id="cmdStart" value="Realizar proximo" onclick="faseStart();" name="cerrar" <% if not flagPuedeCerrar then Response.Write "disabled"%>>
		</td>	
	</tr>
</table>
<table align="center" width="90%" border="1" cellpadding=2 cellspacing=0>	
	<tr>
		<td width="30%" colspan="9" valign="top" align="center" style="height:20;" class="reg_Header_nav"><%=GF_Traducir("Etapas")%></td>
		<td width="70%" rowspan="20" align="center">
			<select size="40"  multiple="multiple" id="tasks" name="tasks" style="width:550pt;scrolling:yes;">
			</select>
		</td>	
	</tr>	
	<!--Etapa 0-->
	<tr>
		<td colspan="9" align="center" id="fase0">
			<div><%=GF_Traducir("Control Previo")%></div>
			<img id="fase0IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>	
	<tr>
		<td colspan="3" align="center" id="fase0a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase0t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase0b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>
	<!--Etapa 1-->
	<tr>
		<td colspan="9" align="center" id="fase1">
			<div><%=GF_Traducir("Inicialización")%></div>
			<img id="fase1IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>		
	<tr>
		<td colspan="3" align="center" id="fase1a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase1t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase1b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>	
	<!--Etapa 2-->
	<tr>
		<td colspan="9" align="center" id="fase2">
			<div><%=GF_Traducir("Stock Fisico")%></div>
			<img id="fase2IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>	
	<tr>
		<td colspan="3" align="center" id="fase2a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase2t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase2b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>
	<!--Etapa 3-->
	<tr>
		<td colspan="9" align="center" id="fase3">
			<div><%=GF_Traducir("Valuación Contable")%></div>

			<img id="fase3IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>		
	<tr>
		<td colspan="3" align="center" id="fase3f" width="33%" bgcolor="">Facturas</td>
		<td colspan="3" align="center" id="fase3r" width="33%" bgcolor="">Vales VRS</td>
		<td colspan="1" align="center" id="fase3t" width="33%" bgcolor="">Transferencias</td>
	</tr>	
	<tr>
		<td align="center" id="fase3fa" bgcolor=""><%=siglasArroyo%></td>
		<td align="center" id="fase3ft" bgcolor=""><%=siglasTransito%></td>
		<td align="center" id="fase3fb" bgcolor=""><%=siglasBahia%></td>
		<td align="center" id="fase3ra" bgcolor=""><%=siglasArroyo%></td>
		<td align="center" id="fase3rt" bgcolor=""><%=siglasTransito%></td>
		<td align="center" id="fase3rb" bgcolor=""><%=siglasBahia%></td>
		<td align="center" id="fase3tt" bgcolor="">Todas</td>
	</tr>
	
	<!--Etapa 4-->
	<tr>
		<td colspan="9" align="center" id="fase4">
			<div><%=GF_Traducir("Aplicación de Precios")%></div>
			<img id="fase4IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>		
	<tr>
		<td colspan="3" align="center" id="fase4a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase4t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase4b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>
	<!--Etapa 5-->
	<tr>
		<td colspan="9" align="center" id="fase5">
			<div><%=GF_Traducir("Armado de Cuenta Corriente")%></div>
			<img id="fase5IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>		
	<tr>
		<td colspan="3" align="center" id="fase5a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase5t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase5b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>
	<!--Etapa 6-->	
	<tr>
		<td colspan="9" align="center" id="fase6">
			<div><%=GF_Traducir("Asientos Contables")%></div>
			<img id="fase6IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>
	<tr>
		<td colspan="3" align="center" id="fase6a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase6t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase6b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>
	<!--Etapa 7-->			
	<tr>
		<td colspan="9" align="center" id="fase7">
			<div><%=GF_Traducir("Finalización")%></div>
			<img id="fase7IMG" style="visibility:hidden;" src="images/icon_ok.gif">
		</td>	
	</tr>	
	<tr>
		<td colspan="3" align="center" id="fase7a" width="33%" bgcolor=""><%=siglasArroyo%></td>
		<td colspan="3" align="center" id="fase7t" width="33%" bgcolor=""><%=siglasTransito%></td>
		<td colspan="3" align="center" id="fase7b" width="33%" bgcolor=""><%=siglasBahia%></td>
	</tr>	
</table>
	<% if isnull(totalDivisionArr) then totalDivisionArr = 0 %>
	<input type="hidden" name="totalDivisionArr" id="totalDivisionArr" value="<%=GF_EDIT_DECIMALS(cDBl(totalDivisionArr),2)%>">
	<% if isnull(totalDivisionTra) then totalDivisionTra = 0 %>
	<input type="hidden" name="totalDivisionTra" id="totalDivisionTra" value="<%=GF_EDIT_DECIMALS(cDBl(totalDivisionTra),2)%>">
	<% if isnull(totalDivisionBba) then totalDivisionBba = 0 %>
	<input type="hidden" name="totalDivisionBba" id="totalDivisionBba" value="<%=GF_EDIT_DECIMALS(cDBl(totalDivisionBba),2)%>">
	<% if isnull(totalDivisionExp) then totalDivisionExp = 0 %>
	<input type="hidden" name="totalDivisionExp" id="totalDivisionExp" value="<%=GF_EDIT_DECIMALS(cDBl(totalDivisionExp),2)%>">

</form>
</body>
</html>