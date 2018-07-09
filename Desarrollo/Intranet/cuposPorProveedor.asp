<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientostitulos.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/cor-IncludePC.asp"-->

<%

locale=session.lcid
session.lcid=2057	'Formato dd/mm/aaaa

Const PERIODO = 9 'es el tamanio de periodo - 1. P ej. si es 13, entonces va a mostrar fechas de dos semanas (14 dias) desde la fecha indicada
Dim arrFechas(9)' si se cambia la const PERIODO, tambien cambiar el valor numerico en la declaracion de array arrFechas

Const CORREDOR_PRIORIDAD = 1
Const VENDEDOR_PRIORIDAD = 2
Const CUPOS_TODOS = 3


'-----------------------------------------------------------------------------------------------
Function obtenerCupos(idProveedor, fecha, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor) 
	Dim strSQL, rs, conn
	Dim fechaInicio, fechaFin
	fechaInicio = fecha
	fechaFin = GF_DTE2FN(dateadd("d",PERIODO,GF_FN2DTE(fecha)))

	strSQL = ""
	if chkEnviados = MOSTRAR_NO_ENVIADOS then strSQL = "select * from("
	strSQL = strSQL & "select cufccp as fechaCupo, cuccor as cdCorredor, cucven as cdVendedor, cuncto as nroContrato, cucpro as cdProducto, cucsuc as cdSucursal, cucope as cdOperacion, cuacos as anioCosecha,  Sum(cucccp "
	if chkEnviados = MOSTRAR_NO_ENVIADOS then strSQL= strSQL & "- case when cupos is null then 0 else cupos end"
	strSQL = strSQL & ") as nroCamiones"
	strSQL = strSQL & " from merfl.mer517f1 " 
	if chkEnviados = MOSTRAR_NO_ENVIADOS then strSQL = strSQL & " left join toepferdb.tblcuposinformados  on CODIGOCUPO=CUCODI and CUCPRO=PRODUCTO and CUCSUC=SUCURSAL and CUCOPE=OPERACION and CUNCTO=NUMERO and CUACOS=COSECHA"
	strSQL = strSQL & " where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin
	strSQL = strSQL & " and (cuccor = " & idProveedor & " or cucven = " & idProveedor & ") "	
	
	
	'aplicarFiltros
	if fltrCdProducto <> "" then strSQL = strSQL & " and cucpro = " & fltrCdProducto
	if fltrCdSucursal <> "" then strSQL = strSQL & " and cucsuc = " & fltrCdSucursal
	if fltrCdOperacion <> "" then strSQL = strSQL & " and cucope = " & fltrCdOperacion
	if fltrNroContrato <> "" then strSQL = strSQL & " and cuncto = " & fltrNroContrato
	if fltrAnioCosecha <> "" then strSQL = strSQL & " and cuacos = " & fltrAnioCosecha
	if fltrPuerto <> 0 then strSQL = strSQL & " and cucdes = " & fltrPuerto
	if fltrCorredor <> 0 then strSQL = strSQL & " and (cuccor = " & fltrCorredor & " or (cucven = " & fltrCorredor & " and cuccor = " & SIN_CORREDOR & "))"
	if fltrVendedor <> 0 then strSQL = strSQL & " and cucven = " & fltrVendedor
	
	strSQL = strSQL & " group by cuncto, cucpro, cucsuc, cucope, cuacos, cufccp, cuccor, cucven "
	if chkEnviados = MOSTRAR_NO_ENVIADOS then strSQL = strSQL & ") as tablaGral where tablaGral.nroCamiones <> 0 "
	strSQL = strSQL & " order by nroContrato, fechaCupo "
	'Response.Write strSQL
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	Set obtenerCupos = rs
End Function
'-----------------------------------------------------------------------------------------------
Function obtenerCorredores(fecha) 
	Dim strSQL, rs, conn
	Dim fechaInicio, fechaFin
	fechaInicio = fecha
	fechaFin = GF_DTE2FN(dateadd("d",PERIODO,GF_FN2DTE(fecha)))

	strSQL = "select cuccor as cdCorredor from ("
	strSQL = strSQL & " select cuccor, sum(cucccp - case when cupos is null then 0 else cupos end ) as nroCamiones "
	strSQL = strSQL & " from merfl.mer517f1 left join toepferdb.tblcuposinformados on codigocupo = cucodi "
	strSQL = strSQL & " where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and cuccor <> " & SIN_CORREDOR
	
	'aplicarFiltros
	if fltrCdProducto <> "" then strSQL = strSQL & " and cucpro = " & fltrCdProducto
	if fltrCdSucursal <> "" then strSQL = strSQL & " and cucsuc = " & fltrCdSucursal
	if fltrCdOperacion <> "" then strSQL = strSQL & " and cucope = " & fltrCdOperacion
	if fltrNroContrato <> "" then strSQL = strSQL & " and cuncto = " & fltrNroContrato
	if fltrAnioCosecha <> "" then strSQL = strSQL & " and cuacos = " & fltrAnioCosecha
	if fltrPuerto <> 0 then strSQL = strSQL & " and cucdes = " & fltrPuerto
	
	strSQL = strSQL & " group by cuncto, cucpro, cucsuc, cucope, cuacos, cufccp, cuccor, cucven) tabla where tabla.nroCamiones <> 0 group by cuccor "
	strSQL = strSQL & " union " 
	strSQL = strSQL & " select cucven as cdCorredor from ("
	strSQL = strSQL & " select cucven, sum(cucccp - case when cupos is null then 0 else cupos end ) as nroCamiones "
	strSQL = strSQL & " from merfl.mer517f1 left join toepferdb.tblcuposinformados on codigocupo = cucodi "
	strSQL = strSQL & " where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and cuccor = " & SIN_CORREDOR
	
	'aplicarFiltros
	if fltrCdProducto <> "" then strSQL = strSQL & " and cucpro = " & fltrCdProducto
	if fltrCdSucursal <> "" then strSQL = strSQL & " and cucsuc = " & fltrCdSucursal
	if fltrCdOperacion <> "" then strSQL = strSQL & " and cucope = " & fltrCdOperacion
	if fltrNroContrato <> "" then strSQL = strSQL & " and cuncto = " & fltrNroContrato
	if fltrAnioCosecha <> "" then strSQL = strSQL & " and cuacos = " & fltrAnioCosecha
	if fltrPuerto <> 0 then strSQL = strSQL & " and cucdes = " & fltrPuerto
	
	strSQL = strSQL & " group by cuncto, cucpro, cucsuc, cucope, cuacos, cufccp, cuccor, cucven) tabla where tabla.nroCamiones <> 0 group by cucven order by cdCorredor"

	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	Set obtenerCorredores = rs
End Function

'-----------------------------------------------------------------------------------------------
Function getTotalCamionesContratoPorFecha(cupos, fechaCorriente)
dim totalCamiones
totalCamiones = 0
cupos.movefirst
while not cupos.eof
	if CLng(cupos("fechaCupo")) = CLng(fechaCorriente) then totalCamiones = totalCamiones + CInt(cupos("nroCamiones"))
cupos.movenext
wend
	getTotalCamionesContratoPorFecha = totalCamiones
end function
'-----------------------------------------------------------------------------------------------
Function getTotalCamiones(cupos)
dim totalCamiones
totalCamiones = 0
cupos.movefirst
while not cupos.eof
	totalCamiones = totalCamiones + CInt(cupos("nroCamiones"))
cupos.movenext
wend
	getTotalCamiones = totalCamiones
end function

'-----------------------------------------------------------------------------------------------
sub armarArrayFechas(fecha, arrFechas)
dim fechaCorriente, contDias
	fechaCorriente = fecha
	contDias = 0
	while contDias <= PERIODO
		arrFechas(contDias) = fechaCorriente
		contDias = contDias + 1
		fechaCorriente = GF_DTEADD(fechaCorriente,1,"D")		
	wend	
end sub
'-----------------------------------------------------------------------------------------------
function determinarSituacion()
	'Caso idProveedor <> 0, fltrCorredor = 0, fltrVendedor = 0 - (inicio de pagina por primera vez con llamado desde as400)
		'En este caso determinar, si Proveedor es Corredor o Vendedor.
		'Setear con este Proveedor fltrCorredor/dsCorredor o fltrVendedor/dsVendedor segun que tipo es Proveedor
		'despues proceder, dependiendo, que tipo de proveedor es segun situacion 2 o 3 o 4
	'Caso idProveedor = 0, fltrCorredor = 0, fltrVendedor = 0 - (sacaron todos los filtros)
		'En este caso hay que mostrar todos los cupos para esta fecha
		'se procede poner variable-flag 
		
	'Situacion 1: (fltrCorredor <> 0 and fltrVendedor = 0) or (fltrCorredor <> 0 and fltrVendedor <> 0)
		'En este caso cargar idProveedor con la variable flrtCorredor y dsCorredor con dsProveedor
		'La columna de la tabla pasara ser de Vendedor		
	'Situacion 2: fltrCorredor = 0, fltrVendedor <> 0
		'En este caso cargar idProveedor con la variable flrtVendedor y dsVendedor con dsProveedor
		'La columna de la tabla pasara ser de Corredor

	if fltrCorredor = 0 and fltrVendedor = 0 and idProveedor <> 0 then 
	
		if GF_ES_CORREDOR(idProveedor) then
			'es corredor
			fltrCorredor = idProveedor
			dsCorredor = dsProveedor
		else
			'es vendedor
			fltrVendedor = idProveedor
			dsVendedor = dsProveedor
		end if
	end if
	if fltrCorredor = 0 and fltrVendedor = 0 and idProveedor = 0 then 
	'Situacion 3
		determinarSituacion	= CUPOS_TODOS
	end if
	if fltrCorredor <> 0 then
		'Situacion 1
		idProveedor = fltrCorredor
		dsProveedor = dsCorredor
		determinarSituacion = CORREDOR_PRIORIDAD
	end if
	if fltrCorredor = 0 and fltrVendedor <> 0 then
		'Situacion 2
		idProveedor = fltrVendedor
		dsProveedor = dsVendedor
		determinarSituacion	= VENDEDOR_PRIORIDAD
	end if	
end function
'-----------------------------------------------------------------------------------------------
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim cupos, rsPuertos, rsCupos, conn, strSQL
Dim idProveedor, usr, dsProveedor
Dim fecha
Dim paginaActual, mostrar, lineasTotales
dim nroContratoAux, cdProductoAux, cdSucursalAux, cdOperacionAux, anioCosechaAux
Dim fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor,dsCorredor, dsVendedor
Dim chkEnviados
dim msg
dim eMails
dim situacionCorredorVendedor
nroContratoAux = 0
GP_ConfigurarMomentos
Dim corredores, vendedores



idProveedor = CLng(GF_PARAMETROS7("id",0,6))
dsProveedor = getDescripcionProveedor(idProveedor)

if (session("Usuario") = "") then session("Usuario") = ucase(Right(Request.ServerVariables("LOGON_USER"),3))
usr = session("Usuario")

fecha = GF_PARAMETROS7("fecha",0,6)
if fecha = 0 then fecha = CLng(Left(session("MmtoDato"),8)) '20100821
msg = GF_PARAMETROS7("msg",0,6)
eMails = GF_PARAMETROS7("mails","",6)
if msg = 1 then call setInfo(MAIL_ENVIO_EXITOSO)
'parametros de filtros
fltrCdProducto = GF_PARAMETROS7("fltrCdProducto","",6)
fltrCdSucursal = GF_PARAMETROS7("fltrCdSucursal","",6)
fltrCdOperacion = GF_PARAMETROS7("fltrCdOperacion","",6)
fltrNroContrato = GF_PARAMETROS7("fltrNroContrato","",6)
fltrAnioCosecha = GF_PARAMETROS7("fltrAnioCosecha","",6)
fltrPuerto = GF_PARAMETROS7("fltrPuerto",0,6)
fltrCorredor = GF_PARAMETROS7("fltrCorredor",0,6)
accion = GF_PARAMETROS7("accion","",6)

dsCorredor = GF_PARAMETROS7("dsCorredor","",6)
fltrVendedor = GF_PARAMETROS7("fltrVendedor",0,6)

dsVendedor = GF_PARAMETROS7("dsVendedor","",6)	
hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
if busquedaActiva = 1 then hayBusqueda = true

chkEnviados = GF_PARAMETROS7("chkEnviados",0,6)
if chkEnviados = 0 then chkEnviados = MOSTRAR_NO_ENVIADOS

situacionCorredorVendedor = determinarSituacion()

if (accion = ACCION_GRABAR) then 	
	Call grabarMails(emails, idProveedor)
end if

Call armarArrayFechas(fecha, arrFechas)

if situacionCorredorVendedor = CUPOS_TODOS then
	Set corredores = obtenerCorredores(fecha)
else
	Set cupos = obtenerCupos(idProveedor, fecha, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor)
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Cupos</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">

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
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">

	var ch = new channel();	
	var arrContratos = new Array();
	var tb = new Toolbar('toolbar', 7, "images/cupos/");
	var idButtonEnviar = 0;
	var ctoSelected = 0;
	var sent=0;
	
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	
	function volver() {
		location.href = "almacenIndex.asp";
	}
	
	function irHome() {
		location.href = "almacenIndex.asp";
	}
	function CerrarCal(cal) {
		cal.hide();
		submitInfo();
	}
	
	function SeleccionarCal(cal, date) {
		var str= new String(date);
		var anio = str.substring(6);
		var mes = str.substring(3,5);
		var dia = str.substring(0,2);
		var fechaHoy = new Date();
		var fecha = new Date(date);
		document.getElementById("issuedateDiv").innerHTML = str;
		document.getElementById("fecha").value = parseInt(anio + mes + dia);	
		if (cal) CerrarCal(cal);			
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
	function submitInfo() {
		document.getElementById("frmSel").submit();
	}
	function mostrarNoEnviados(chk) {
		if (chk == <%=MOSTRAR_NO_ENVIADOS%>) {
			document.getElementById("chkEnviados").value = <%=MOSTRAR_ENVIADOS%>;
		}else{
			document.getElementById("chkEnviados").value = <%=MOSTRAR_NO_ENVIADOS%>;
		}
		submitInfo();
	}
	
	function enviarGrupo(idProveedor, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor, chkEnviados,periodo) {
		if (ctoSelected == 0) {
			enviarTodo(idProveedor, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor, chkEnviados,periodo);
		} else {
			enviarSeleccionados(idProveedor, chkEnviados);
		}
		
	}
	function enviarTodo(idProveedor, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor, chkEnviados,periodo) {		
			var filtros = '&fltrCdProducto=' + fltrCdProducto + '&fltrCdSucursal=' + fltrCdSucursal + '&fltrCdOperacion=' + fltrCdOperacion;
			filtros = filtros + '&fltrNroContrato=' + fltrNroContrato + '&fltrAnioCosecha=' + fltrAnioCosecha + '&fltrPuerto=' + fltrPuerto;
			filtros = filtros + '&fltrCorredor=' + fltrCorredor + '&fltrVendedor=' + fltrVendedor + "&chkEnviados=" + chkEnviados;		
			var puw = new PopUpWindow('popupEnviarMail','cuposGetDescripcionMail.asp?id=' + idProveedor + '&fecha=<%=fecha%>&periodo=<%=PERIODO%>&pdf_accion=<%=PDF_FILE_MODE %>&usr=<%=usr%>' + filtros,'600','350','Enviar Mail');
			puw.onHideEnd = 'submitInfo()';			
	}
	

	function mailSeleccion_callback() {
		sent -= 1;
		if (sent == 0) submitInfo();		
	}	
	
	function enviarSeleccionados(idProveedor, chkEnviados) {					
		tb.changeLook(idButtonEnviar, "Loading2.gif", 'Enviar Seleccion');	
		sent = 0;
		for (theKey in arrContratos) {
			if (arrContratos[theKey] == 'V') {
				fields = theKey.split('|');	
				var params = "&nroContrato=" + fields[0] + "&cdProducto=" + fields[1] + "&cdSucursal=" + fields[2] + "&cdOperacion=" + fields[3] + "&anioCosecha=" + fields[4] + "&chkEnviados=" + chkEnviados;
				ch.bind('cuposPorProveedorPrintXLS.asp?id=' + idProveedor + '&fecha=<%=fecha%>&periodo=<%=PERIODO%>&usr=<%=usr%>&mostrarConfirmacion=N&xls_accion=<%=XLS_FILE_MODE%>' + params, "mailSeleccion_callback()");
				ch.send();
				sent += 1;
				arrContratos[theKey] = 'F'
			}
		}				
	}
		
	function marcarContrato(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, chkObj) {
		var arrKey = nroContrato + "|" + cdProducto + "|" + cdSucursal + "|" + cdOperacion + "|" + anioCosecha;
		if (chkObj.checked) {
			arrContratos[arrKey] = 'V';									
			tb.changeLook(idButtonEnviar, "Mail-16x16.png", 'Enviar Seleccion');
			ctoSelected += 1;
		} else {			
			arrContratos[arrKey] = 'F';			
			ctoSelected -= 1;
		}		
		if (ctoSelected == 0) tb.changeLook(idButtonEnviar, "Mail-16x16.png", 'Enviar Todo');
	}
	
	function enviarContrato(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, chkEnviados) {
		var params = "&nroContrato=" + nroContrato + "&cdProducto=" + cdProducto + "&cdSucursal=" + cdSucursal + "&cdOperacion=" + cdOperacion + "&anioCosecha=" + anioCosecha + "&chkEnviados=" + chkEnviados;
		var puw = new PopUpWindow('popupEnviarMail','cuposGetDescripcionMail.asp?id=<%=idProveedor%>&fecha=<%=fecha%>&periodo=<%=PERIODO%>&pdf_accion=<%=PDF_FILE_MODE %>&usr=<%=usr%>' + params,'600','350','Enviar Mail');	
		puw.onHideEnd = 'submitInfo()';		
	}
	function verTodo(idProveedor, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor, chkEnviados, periodo) {		
		var filtros = '&fltrCdProducto=' + fltrCdProducto + '&fltrCdSucursal=' + fltrCdSucursal + '&fltrCdOperacion=' + fltrCdOperacion;
		filtros = filtros + '&fltrNroContrato=' + fltrNroContrato + '&fltrAnioCosecha=' + fltrAnioCosecha + '&fltrPuerto=' + fltrPuerto;
		filtros = filtros + '&fltrCorredor=' + fltrCorredor + '&fltrVendedor=' + fltrVendedor + '&chkEnviados=' + chkEnviados+ "&periodo=" + periodo;
		window.open("cuposPorProveedorPrint.asp?id=" + idProveedor + "&fecha=<%=fecha%>&periodo=<%=PERIODO%>&pdf_accion=<%=PDF_STREAM_MODE %>" + filtros);
	}
	function verTodoXLS(idProveedor, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor, chkEnviados, periodo) {		
		var filtros = '&fltrCdProducto=' + fltrCdProducto + '&fltrCdSucursal=' + fltrCdSucursal + '&fltrCdOperacion=' + fltrCdOperacion;
		filtros = filtros + '&fltrNroContrato=' + fltrNroContrato + '&fltrAnioCosecha=' + fltrAnioCosecha + '&fltrPuerto=' + fltrPuerto;
		filtros = filtros + '&fltrCorredor=' + fltrCorredor + '&fltrVendedor=' + fltrVendedor + '&chkEnviados=' + chkEnviados+ "&periodo=" + periodo;
		window.open("cuposPorProveedorPrintXLS.asp?id=" + idProveedor + "&fecha=<%=fecha%>&periodo=<%=PERIODO%>&xls_accion=<%=XLS_STREAM_MODE %>" + filtros);
	}
	function verContrato(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, chkEnviados) {
		var params = "&nroContrato=" + nroContrato + "&cdProducto=" + cdProducto + "&cdSucursal=" + cdSucursal + "&cdOperacion=" + cdOperacion + "&anioCosecha=" + anioCosecha + "&chkEnviados=" + chkEnviados;
		window.open("cuposPorProveedorPrint.asp?id=<%=idProveedor%>&fecha=<%=fecha%>&periodo=<%=PERIODO%>&pdf_accion=<%=PDF_STREAM_MODE%>" + params); 
	}
	function verContratoXLS(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, chkEnviados) {
		var params = "&nroContrato=" + nroContrato + "&cdProducto=" + cdProducto + "&cdSucursal=" + cdSucursal + "&cdOperacion=" + cdOperacion + "&anioCosecha=" + anioCosecha + "&chkEnviados=" + chkEnviados;
		window.open("cuposPorProveedorPrintXLS.asp?id=<%=idProveedor%>&fecha=<%=fecha%>&periodo=<%=PERIODO%>&xls_accion=<%=XLS_STREAM_MODE%>" + params); 
	}
	
	function validarEmail(valor) {
		var filter=/^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i		
		if (!filter.test(valor)){		
				alert("La dirección de email " + valor + " es incorrecta.\n" + "Ingrese los mails separados por punto y coma (;)");
				return false;
			}
		return true;		
	}	
	function agregarMail(){
		var strMails = document.getElementById('mails').value;
		strMails = strMails.replace(/\n/gi, "");
		var arrayMails = strMails.split(";");
		if (arrayMails.length < 10 ){
			var mailsCorrectos = true;
			for (i=0;i<arrayMails.length;i++){	
				if (!validarEmail(arrayMails[i].toString().toLowerCase())) {
					mailsCorrectos = false;
				}
			}
			if (mailsCorrectos){
				document.getElementById("accion").value= '<% =ACCION_GRABAR %>';
				submitInfo();			
			}
		}else{
			alert("solo se permite ingresar hasta diez e-mails por proveedor");
		}		
	}
	function bodyOnLoad() {			
		tb.addButton("Home-16x16.png", "Home", "irHome()");	
		tb.addButtonREFRESH("Recargar", "submitInfo()");
		
		<%	if (situacionCorredorVendedor <> CUPOS_TODOS) then 
				if (idProveedor <> MERCADO_A_TERMINO) then
		%>
			idButtonEnviar = tb.addButton("Mail-16x16.png", "Enviar Todo", "enviarGrupo('<%=idProveedor%>', '<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<%=fltrPuerto%>','<%=fltrCorredor%>','<%=fltrVendedor%>', '<%=chkEnviados%>', '<%=PERIODO%>')");	
			tb.addButton("pdf.gif", "Ver Todo", "verTodo('<%=idProveedor%>','<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<%=fltrPuerto%>','<%=fltrCorredor%>','<%=fltrVendedor%>', '<%=chkEnviados%>', '<%=PERIODO%>')");
			tb.addButton("excel.gif", "Ver Todo", "verTodoXLS('<%=idProveedor%>','<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<%=fltrPuerto%>','<%=fltrCorredor%>','<%=fltrVendedor%>', '<%=chkEnviados%>', '<%=PERIODO%>')");												
		<%		end if			%>
		<% else %>
			idButtonEnviar = tb.addButton("Mail-16x16.png", "Enviar Todo", "enviarEnBatch('<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<%=fltrPuerto%>')");			
		<% end if%>
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");				
		tb.draw();
			<%	if (hayBusqueda) then %>
					tb.changeState(swt);
					<%if situacionCorredorVendedor <> CUPOS_TODOS then%>			
						startMagicSearch();
					<%end if%>
			<%end if%>
		pngfix();
	}
	
<%if situacionCorredorVendedor = CUPOS_TODOS then
	if not corredores.eof then%>	
	function enviarEnBatch(fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto){
	<% if not corredores.eof then
			if (CInt(corredores("cdCorredor")) = MERCADO_A_TERMINO) then corredores.movenext
			'Si se cambió puede ser fin de archivo. Se verifica nuevamente.
			if not corredores.eof then
		%>			
			envio<%=CInt(corredores("cdCorredor"))%>(fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto);
		<%	end if
		end if%>
	}
	
	//crear funciones javascript AJAX para envio en batch de los corredores
	<%while not corredores.eof%>
		function envio<%=CInt(corredores("cdCorredor"))%>(fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto){
			indicarProcesoEnvio(<%=CInt(corredores("cdCorredor"))%>);
			var filtros = '&fltrCdProducto=' + fltrCdProducto + '&fltrCdSucursal=' + fltrCdSucursal + '&fltrCdOperacion=' + fltrCdOperacion;
			filtros = filtros + '&fltrNroContrato=' + fltrNroContrato + '&fltrAnioCosecha=' + fltrAnioCosecha + '&fltrPuerto=' + fltrPuerto;
			ch.bind('cuposPorProveedorPrintXLS.asp?id=<%=CInt(corredores("cdCorredor"))%>&fecha=<%=fecha%>&periodo=<%=PERIODO%>&mostrarConfirmacion=N&chkEnviados=<%=MOSTRAR_NO_ENVIADOS%>&xls_accion=<%=XLS_FILE_MODE%>' + filtros, "envio<%=CInt(corredores("cdCorredor"))%>_Callback()");
			ch.send();
		}
		
		function envio<%=CInt(corredores("cdCorredor"))%>_Callback(){
			if(ch.response() != ''){
				avisarMailAusente(<%=CInt(corredores("cdCorredor"))%>)
			}else{
				cerrarProcesoEnvio(<%=CInt(corredores("cdCorredor"))%>);
			};
			
	<%
		corredores.movenext
			if not corredores.eof then
				if (CInt(corredores("cdCorredor")) = MERCADO_A_TERMINO) then corredores.movenext
				'Si se cambió puede ser fin de archivo. SE verifica nuevamente.
				if not corredores.eof then
			
	%>
			envio<%=CInt(corredores("cdCorredor"))%>('<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<%=fltrPuerto%>');			
	<%			end if
			end if
	%>
		}
	<%wend
	corredores.movefirst
	%>
	
	function indicarProcesoEnvio(idProveedor){
		document.getElementById("prov" + idProveedor).innerHTML = "<img src='images/cupos/Loading2.gif'></img>";
	}
	
	function cerrarProcesoEnvio(idProveedor){
		document.getElementById("prov" + idProveedor).innerHTML = "<img src='images/cupos/icon_ok.gif'></img>";
	}
	function avisarMailAusente(idProveedor){
		document.getElementById("prov" + idProveedor).innerHTML = "<img src='images/cupos/button_cancel.png'></img>";
	}
<%end if%>
<%end if%>
	
	function startMagicSearch() {		
		var msCorredor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas&linea=0");
		msCorredor.setMinChar(3);
		msCorredor.setToken(";");
		msCorredor.onBlur = seleccionarCorredor;
		msCorredor.setValue('<%=dsCorredor%>');
		
		var msVendedor = new MagicSearch("", "companyName1", 30, 2, "comprasStreamElementos.asp?tipo=empresas&linea=1");
		msVendedor.setMinChar(3);
		msVendedor.setToken(";");
		msVendedor.onBlur = seleccionarVendedor;
		msVendedor.setValue('<%=dsVendedor%>');
		
		
	}
	function seleccionarCorredor(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("fltrCorredor").value = arr[0];
			document.getElementById("dsCorredor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("fltrCorredor").value = 0;
			if (desc == "") document.getElementById("dsCorredor").value = "";
			ms.setValue("");
		}
	}
	function seleccionarVendedor(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("fltrVendedor").value = arr[0];
			document.getElementById("dsVendedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("fltrVendedor").value = 0;
			if (desc == "") document.getElementById("dsVendedor").value = "";
			ms.setValue("");
		}
	}
	
	function buscarOn() {
		document.getElementById("busqueda").className = "";
		document.getElementById("busquedaActiva").value = "1";
		<%if situacionCorredorVendedor <> CUPOS_TODOS then%>	
			startMagicSearch();
		<%end if%>
	}
	function buscarOff() {
		document.getElementById("busqueda").className = "divOculto";
		document.getElementById("busquedaActiva").value = "0";
	}	
	
	function redirigir(id){
	
		location.href = "cuposPorProveedor.asp?id=" + id + "&fecha=<%=fecha%>&periodo=<%=PERIODO%>";
	}
	
</script>
</head>

<%if situacionCorredorVendedor = CUPOS_TODOS then%>
	<body onLoad="bodyOnLoad()">
		<%call GF_TITULO2("kogge64.gif","Corredores y Vendedores con cupos pendientes de envio") %>		
		<div id="toolbar"></div>
		<br>
		<form name="frmSel" id="frmSel" action="cuposPorProveedor.asp" method="post">
		<div id="busqueda" class="divOculto">
		<table width="40%" cellspacing="0" cellpadding="0" align="center" border="0">
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
									<td align="right"><% =GF_TRADUCIR("Contrato") %>:</td>
									<td>
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrCdProducto %>" name="fltrCdProducto" id="fltrCdProducto" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="2" maxLength="1" value="<% =fltrCdSucursal %>" name="fltrCdSucursal" id="fltrCdSucursal" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrCdOperacion %>" name="fltrCdOperacion" id="fltrCdOperacion" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="5" maxLength="5" value="<% =fltrNroContrato %>" name="fltrNroContrato" id="fltrNroContrato" onKeyPress="return controlIngreso (this, event, 'N');"> /
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrAnioCosecha %>" name="fltrAnioCosecha" id="fltrAnioCosecha" onKeyPress="return controlIngreso (this, event, 'N');">                                    
	                                
	                          		</td>  
								</tr>

	                            <tr>
									<td align="right"><% =GF_TRADUCIR("Puerto") %>:</td>
									<%
									Set rsPuertos = obtenerListaPuertos()%>   
	                                <td>                                
										<select id="fltrPuerto" name="fltrPuerto">
											<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -
											<%	
											while (not rsPuertos.eof)	%>
												<option value="<% =rsPuertos("IDPUERTO") %>" <% if (CInt(rsPuertos("IDPUERTO")) = CInt(fltrPuerto)) then response.write "selected='true'" %>><% =rsPuertos("IDPUERTO") %> - <% =GF_TRADUCIR(rsPuertos("DSPUERTO")) %>
												<%		
												rsPuertos.MoveNext()
											wend 	
											%>		
										</select>		
	                                </td>								
	                            </tr>									
	                            <tr>
									<td colspan="4" align="center"><input type="button" value="Buscar..." id=submit1 name=submit1 onclick='submitInfo();'></td>						
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
		<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">
		<input type="hidden" name="usr" id="usr" value="<% =usr %>">
		<br>
		<table class="reg_Header" id="TAB3" align="center" border="0">
			<tr>
				<td class="reg_Header_navdos">
				<% =GF_TRADUCIR("Fecha") %>
				</td>
				<td align="center" >
					<div id="issuedateDiv"><% =GF_FN2DTE(fecha) %></div>															
					<input type="hidden" id="fecha" name="fecha" value="<% =fecha %>">
				</td>
				<td align="center" >
					<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCal)"><img id="imgEmision" src="images/DATE.gif"></a>
				</td>
		</tr>
		</table>	
		<br>
		<table align="center" class="reg_Header">
					<tr class="reg_Header_nav">
						<td  style="text-align: center" colspan="2"><% =GF_TRADUCIR("Proveedores") %></td>
						<td style="text-align: center"><% =GF_TRADUCIR("Envio") %></td>
						<td style="text-align: center"><% =GF_TRADUCIR("PDF") %></td>
						<td style="text-align: center"><% =GF_TRADUCIR("Excel") %></td>
						<td style="text-align: center"><% =GF_TRADUCIR("Enviar") %></td>
											
					</tr>		
		<%	if (not corredores.eof) then			
				while ((not corredores.eof))%>
					<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)" >
						<td align="center" onclick="redirigir(<%=corredores("cdCorredor")%>);">	<% =CInt(corredores("cdCorredor"))%></td>
						<td onclick="redirigir(<%=corredores("cdCorredor")%>);">&nbsp;	<% =getDescripcionProveedor(CInt(corredores("cdCorredor")))%></td>	
						<td align="center" ><div id="prov<%=corredores("cdCorredor")%>">&nbsp;</div></td>
						<% if (CInt(corredores("cdCorredor")) = MERCADO_A_TERMINO) then %>
						<td class="reg_Header_nav" align="center" ></td>
						<td class="reg_Header_nav" align="center" ></td>
						<td class="reg_Header_nav" align="center" ></td>
						<% else %>
						<td class="reg_Header_nav" align="center" ><img onclick="javascript:verTodo('<%=CInt(corredores("cdCorredor"))%>', '<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<% =fltrPuerto %>','','', '<%=MOSTRAR_NO_ENVIADOS%>', '<%=PERIODO%>');" src="images/cupos/pdf.gif" ></td>
						<td class="reg_Header_nav" align="center" ><img onclick="javascript:verTodoXLS('<%=CInt(corredores("cdCorredor"))%>', '<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>','<% =fltrPuerto %>','','', '<%=MOSTRAR_NO_ENVIADOS%>', '<%=PERIODO%>');" src="images/cupos/excel.gif" ></td>					
						<td class="reg_Header_nav" align="center" ><img onclick="javascript:enviarTodo('<%=CInt(corredores("cdCorredor"))%>', '<%=fltrCdProducto%>', '<%=fltrCdSucursal%>', '<%=fltrCdOperacion%>','<%=fltrNroContrato%>','<%=fltrAnioCosecha%>',' <% =fltrPuerto %>','','', '<%=MOSTRAR_NO_ENVIADOS%>', '<%=PERIODO%>');"  src="images/cupos/Mail-16x16.png"></td>
						<%end if%>
						
					</tr>
				<%corredores.movenext		
				wend		
			else%>
				<tr class="TDNOHAY"><td><% =GF_TRADUCIR("No hay informacion de cupos de los corredores") %></td></tr>		
			<%end if%>					
			</table>		
	</form>	
	</body>
	</html>
<%else%>
	<body onLoad="bodyOnLoad()">
		<%call GF_TITULO2("kogge64.gif","Cupos Por Contratos") %>		
		<div id="toolbar"></div>
		<br>
		<form name="frmSel" id="frmSel" action="cuposPorProveedor.asp" method="post">
		<div id="busqueda" class="divOculto">
		<table width="40%" cellspacing="0" cellpadding="0" align="center" border="0">
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
									<td align="right"><% =GF_TRADUCIR("Contrato") %>:</td>
									<td>
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrCdProducto %>" name="fltrCdProducto" id="fltrCdProducto" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="2" maxLength="1" value="<% =fltrCdSucursal %>" name="fltrCdSucursal" id="fltrCdSucursal" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrCdOperacion %>" name="fltrCdOperacion" id="fltrCdOperacion" onKeyPress="return controlIngreso (this, event, 'N');"> -
	                                    <input type="text" size="5" maxLength="5" value="<% =fltrNroContrato %>" name="fltrNroContrato" id="fltrNroContrato" onKeyPress="return controlIngreso (this, event, 'N');"> /
	                                    <input type="text" size="2" maxLength="2" value="<% =fltrAnioCosecha %>" name="fltrAnioCosecha" id="fltrAnioCosecha" onKeyPress="return controlIngreso (this, event, 'N');">                                    
	                                
	                          		</td>  
								</tr>

	                            <tr>
									<td align="right"><% =GF_TRADUCIR("Puerto") %>:</td>
									<%
									Set rsPuertos = obtenerListaPuertos()%>   
	                                <td>                                
										<select id="fltrPuerto" name="fltrPuerto">
											<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -
											<%	
											while (not rsPuertos.eof)	%>
												<option value="<% =rsPuertos("IDPUERTO") %>" <% if (CInt(rsPuertos("IDPUERTO")) = CInt(fltrPuerto)) then response.write "selected='true'" %>><% =rsPuertos("IDPUERTO") %> - <% =GF_TRADUCIR(rsPuertos("DSPUERTO")) %>
												<%		
												rsPuertos.MoveNext()
											wend 	
											%>		
										</select>		
	                                </td>								
	                            </tr>									

								<tr>
									<td align="right"><% = GF_TRADUCIR("Corredor") %>:</td>
									<td>
										<div id="companyName0"></div>			
										<input type="hidden" id="fltrCorredor" name="fltrCorredor" value="<% =fltrCorredor %>">
										<input type="hidden" id="dsCorredor" name="dsCorredor" value="<% =dsCorredor %>">
									</td>								
	                                                              
	                            </tr>
	                            <tr>
									<td align="right"><% = GF_TRADUCIR("Vendedor") %>:</td>
									<td>
										<div id="companyName1"></div>			
										<input type="hidden" id="fltrVendedor" name="fltrVendedor" value="<% =fltrVendedor %>">
										<input type="hidden" id="dsVendedor" name="dsVendedor" value="<% =dsVendedor %>">
									</td>
	                            </tr>
	                            <tr>
									<td colspan="4" align="center"><input type="button" value="Buscar..." id=submit1 name=submit1 onclick='submitInfo();'></td>						
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
		<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">

		<input type="hidden" name="usr" id="usr" value="<% =usr %>">
		<div>
		<br><br>
		<table class="reg_Header" id="TAB3" align="center" border="0">	
		<tr>
			<td class="reg_Header_navdos">
				<% if situacionCorredorVendedor = CORREDOR_PRIORIDAD then%>
					<% =GF_TRADUCIR("Corredor") %>
				<%else %>
					<% =GF_TRADUCIR("Vendedor") %>
				<%end if%>
			</td>
			<td align="center" colspan="2">
				<% =dsProveedor %>
				</div>																		
			</td>		
		</tr>		
			<%'dibujar los mails del proveedor
				'MERCADO A TERMNIO NO DEBE RECIBIR MAILS!!! SOLO EL CORREDOR (VER MER311FH)
				if (CLng(idProveedor) <> MERCADO_A_TERMINO) then
					dim cantMails, mails(10)
					cantMails = obtenerMailTipo(idProveedor, MAIL_CUPO, mails)
					if cantMails = 0 then%>
						<tr class="reg_Header_navdos"><td colspan="3" class="TDERROR" >
								<% =GF_TRADUCIR("ATENCION") %>:<%=GF_TRADUCIR("Proveedor no tiene mail correspondiente")%>
							</td></tr>
					<%end if%>
						<tr><td class="reg_Header_navdos">
							<% =GF_TRADUCIR("Mails") %>
						</td>				
						<td align="center">
							<textarea id="mails" name="mails" style="text-align: left" wrap="soft" rows="2" cols="30"><% =getStringMailsProveedor(idProveedor) %></textarea>
						</td>
						<td align="center">
							<img src="images/cupos/Guardar.gif" onclick="javascript:agregarMail();"style="cursor:pointer" title="<%=GF_Traducir("Agregar Mail")%>"></img>
						</td>				
						</tr>			
			<%	end if %>
		<tr>
			<td class="reg_Header_navdos">
				<% =GF_TRADUCIR("Fecha") %>
			</td>
			<td align="center" >
				<div id="issuedateDiv"><% =GF_FN2DTE(fecha) %></div>															
				<input type="hidden" id="fecha" name="fecha" value="<% =fecha %>">
			</td>
			<td align="center" >
				<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCal)"><img id="imgEmision" src="images/DATE.gif"></a>
			</td>
		</tr>
		</table>
		<br>
		<% if msg = 1 then %>
			<table align="center" width="90%">
				<tr><td><%=showErrors()%></td></tr>
			</table>
		<%
		msg = 0
		end if%>	
		<br>
		<table align="center" width="90%" class="reg_Header">
			<tr class="reg_Header_nav">
					<td  style="text-align: right" colspan="<%=(PERIODO+8)%>"><% =GF_TRADUCIR("Mostrar solo no enviados") %>&nbsp;<input type='checkbox' <%if chkEnviados = MOSTRAR_NO_ENVIADOS then Response.Write " checked='checked' " %> onclick="mostrarNoEnviados(<%=chkEnviados%>)" /></td>				
					<input type="hidden" name='chkEnviados' id='chkEnviados' value='<%=chkEnviados%>'/>
				</tr>
				<tr class="reg_Header_nav">
					<td  style="text-align: center"><% =GF_TRADUCIR("Contrato") %></td>
					<td  style="text-align: center">
					<% if situacionCorredorVendedor = CORREDOR_PRIORIDAD then%>
						<% =GF_TRADUCIR("Vendedor") %>
					<%else%>
						<% =GF_TRADUCIR("Corredor") %>
					<%end if%>				
					</td>
					<%for i= 0 to PERIODO%>				
					<td  style="text-align: center"><% =left(GF_FN2DTE(arrFechas(i)),5) %></td>
					<%next%>
					<td style="text-align: center"><% =GF_TRADUCIR("Totales") %></td>
					<td style="text-align: center">.</td>
					<td style="text-align: center"><% =GF_TRADUCIR("PDF") %></td>
					<td style="text-align: center"><% =GF_TRADUCIR("XLS") %></td>
					<td style="text-align: center"><% =GF_TRADUCIR("Enviar") %></td>
					
				</tr>		
	<%	nroContratoAux = 0
		index=0
		dim totalCuposCto		
		if (not cupos.eof) then					
		    while (not cupos.eof)   
                nroContratoAux = CLng(cupos("nroContrato"))
			    cdProductoAux = CInt(cupos("cdProducto"))
			    cdSucursalAux = CInt(cupos("cdSucursal"))
			    cdOperacionAux = CInt(cupos("cdOperacion"))
			    anioCosechaAux = CInt(cupos("anioCosecha"))					    
			    totalCuposCto = 0		
			    index=index+1
		        totalCuposCto = 0		    
%>
		        <tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">	
				    <td align="center"><% =GF_EDIT_CONTRATO(cupos("cdProducto"), cupos("cdSucursal"),cupos("cdOperacion"), cupos("nroContrato"), cupos("anioCosecha")) %></td>	
				    <td align="center">
					    <% if situacionCorredorVendedor = CORREDOR_PRIORIDAD then%>
						    <% =getDescripcionProveedor(CInt(cupos("cdVendedor")))%>
					    <%else%>
						    <% =getDescripcionProveedor(CInt(cupos("cdCorredor")))%>						
					    <%end if%>					
				    </td>	
<%					    
                iCelda = 0
			    while ((not cupos.eof) and (iCelda <= PERIODO))				    
			        if ((CLng(arrFechas(iCelda)) = CLng(cupos("fechaCupo"))) and (CLng(nroContratoAux) = CLng(cupos("nroContrato")))) then
			            totalCuposCto = totalCuposCto + CInt(cupos("nroCamiones"))
					    %>
					        <td align="center" ><% =cupos("nroCamiones") %></td>
				        <%  				        
			            cupos.MoveNext()
			        else
			            %>
			                <td align="center" >-</td>
			            <%
			        end if
			        iCelda = iCelda + 1
			    wend
			    'Completo las celdas de la ultima linea.(Salio del ciclo por eof)
			    while iCelda <= PERIODO %>
			        <td align="center" >-</td>
		        <%
			        iCelda = iCelda + 1
		        wend		
		        %>
		            <td class="reg_Header_nav" align="center" ><%=totalCuposCto%></td>
		            <% if (CLng(idProveedor) <> MERCADO_A_TERMINO) then %>
		            <td  class="reg_Header_nav" align="center" ><input type="checkbox" id="Checkbox1" name="chk<% = index %>" onClick="marcarContrato(<%=nroContratoAux%>, <%=cdProductoAux%>, <%=cdSucursalAux%>, <%=cdOperacionAux%>, <%=anioCosechaAux%>, this);"></td>	
		            <% else %>
		            <td  class="reg_Header_nav"></td>
		            <% end if %>
		            <td class="reg_Header_nav" align="center" ><img onclick="javascript:verContrato(<%=nroContratoAux%>, <%=cdProductoAux%>, <%=cdSucursalAux%>, <%=cdOperacionAux%>, <%=anioCosechaAux%>, <%=chkEnviados%>);" src="images/cupos/pdf.gif"></td>								
		            <td class="reg_Header_nav" align="center" ><img onclick="javascript:verContratoXLS(<%=nroContratoAux%>, <%=cdProductoAux%>, <%=cdSucursalAux%>, <%=cdOperacionAux%>, <%=anioCosechaAux%>, <%=chkEnviados%>);" src="images/cupos/excel.gif"></td>								
		            <td class="reg_Header_nav" align="center" ><img onclick="javascript:enviarContrato(<%=nroContratoAux%>, <%=cdProductoAux%>, <%=cdSucursalAux%>, <%=cdOperacionAux%>, <%=anioCosechaAux%>, <%=chkEnviados%>);" src="images/cupos/Mail-16x16.png"></td>								
	            </tr>	
		        <%
            wend			    
%>        
		<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">
			<td class="reg_Header_nav" align="center" colspan="2" ><%=GF_TRADUCIR("Totales")%></td>
			<%for i= 0 to PERIODO%>				
					<td class="reg_Header_nav"  style="text-align: center"><% =getTotalCamionesContratoPorFecha(cupos, arrFechas(i)) %></td>
			<%next%>		
			<td class="reg_Header_nav" align="center" ><%=getTotalCamiones(cupos)%></td>
		</tr>
		<%else%>
			<tr class="TDNOHAY"><td colspan="<%=PERIODO+3%>"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
		<%end if%>			
			</table>
	</form>	
	</body>
	</html>
<%end if
 session.lcid=locale 'volver al formato de fecha original del servidor
%>