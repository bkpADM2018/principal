<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call comprasControlAccesoCM(RES_CC)

dim idPedido, c1, c2, c3, c4, nroLinea, myComentarios, myStyleVisibility, responsableCd, responsableDs
dim index, auxSelected, myAccion, mySeleccion, myResponsable, myNoCotizaState, myNoCotizaValue, esExpo
dim ITproveedor, ITcaracteristica, ITimporte, ITmoneda, ITcondPago, ITfecEntrega, ITnroLinea, ITchkFecha, rsFirmas
dim myQuitar, pct_Abierto, member1, member2, member1Cd, member2Cd, fechaLabel,myImporteSeleccion, myMonedaSeleccion,myPoliza,tienePoliza,auxMoneda
dim saldoPDCCargado, saldoPDCPendiente, saldoPDCTotal, isPendiente, monedaPDC, IdPdc, pcpType, idSector

'-----------------------------------------------------------------------------------------------
Function controlarPlanilla(pCdResponsable, pCdMiembro1, pCdMiembro2, pITfecEntrega, pITimporte, pITmoneda,pProvSel, pMontoPoliza, pITproveedor, pIdDivision, pIdObra, pIdArea,pIdDetalle)
	Dim k, ret, kr, ds, kr1, kr2, kr3
	'Controlo los datos de los proveedores
	ret = true
	if (controlPartidaPresupuestariaPCP(pIdObra, pIdDetalle, pIdArea)) then
	    for k = 1 to nroLinea
		    'Controlo el importe
		    if ((pITimporte(k) = 0) and (pITfecEntrega(k) <> ACCION_PCT_RETIRARSE)) then ret = false    		
		    if(pITproveedor(k) = pProvSel)then
    		    if(Cdbl(pITimporte(k)) < (pMontoPoliza/100))then Call setError(PDC_MONTO_INCORRECTO)
    		    pct_idProveedorElegido = pProvSel    		
    		    if(checkPCPrequiredAFE(idPedido,pITimporte(k) ,pITmoneda(k))) then Call setWarning(PCP_NECESITA_AFE)
		    end if
	    next
	    if not hayError() then
		    if (ret) then
			    ret =false
                if (pIdObra = OBRA_GEID) then
                    'Si selecciono la partida GEID, debe tener completo el campo comentarios
                    if (Trim(myComentarios) = "") then Call setError(SM_OBS_REQUERIDAS)
                end if
			    if (pProvSel > 0) then
				    if (getUserDescription(pCdResponsable) <> "") then
					    'if ((getUserDescription(pCdMiembro1) <> "") or (getUserDescription(pCdMiembro2) <> "") or (getUserDescription(pCdMiembro3) <> "")) then
						    'if (controlResponsables(pCdResponsable, pCdMiembro1, pCdMiembro2, pCdMiembro3, pIdDivision)) then 
							    ret = true							
						    'end if
					    'else					
					    '	Call setError(FALTA_MIEMBROS_ADJUDICACION)
					    'end if
				    else				
					    Call setError(FALTA_RESPONSABLE)
				    end if
			    else			
				    Call setError(FALTA_PROV_ADJUDICADO)
			    end if
		    else		
			    Call setError(IMPORTE_NO_EXISTE)
		    end if
	    else
		    'Monto PDC incorrecto
		    ret = false	
	    end if	
	end if
	controlarPlanilla = ret
End Function

'-----------------------------------------------------------------------------------------------
sub RedimVarialbles(pCant)
	redim ITproveedor(pCant)
	redim ITcaracteristica(pCant)
	redim ITimporte(pCant)
	redim ITmoneda(pCant)
	redim ITcondPago(pCant)
	redim ITfecEntrega(pCant)
	redim ITnroLinea(pCant)
	redim ITchkFecha(pCant)
end sub
'-----------------------------------------------------------------------------------------------
function CotizacionIsOpen(pIdCotizacion, pIdProveedor)
dim rtrn
if pct_tipoCompra <> "T" then 
	rtrn = true
else
	rtrn = false
	strSQL="SELECT FECHAAPERTURA from TBLPCTCOTIZACIONES where IDPEDIDO=" & pIdCotizacion & " and IDPROVEEDOR=" & pIdProveedor & " and not fechaApertura is null"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = true
end if
CotizacionIsOpen = rtrn
end function
'-----------------------------------------------------------------------------------------------
'Se envía al departamento de impuestos una notificación para que autoricen al proveedor.
Function notifySuppliersDept(idProv)
	Dim asunto, msg
	Dim ds, cuit
	
	asunto="Proveedor necesita autorizacion"
	
	ds  = getDescripcionProveedor(idProv)
	cuit= getCUITProveedor(idProv)
	
	msg = "Para el pedido de precio" & pct_cdPedido & " - " & pct_titulo & " cuyo solicitante es " & pct_dsSolicitante & ", se ha decidido adjudicar al proveedor " & ds & " (CUIT: " & cuit & "), el mismo no esta registrado en nuestro maestro de proveedores y se requiere completar este paso para concluir el proceso de adjudicación."
	Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web -" & asunto) & ": " & pct_cdPedido, msg, SENDER_LICITACIONES, SENDER_SUPPLIERS)
	
End Function
'----------------------------------------------------------------------------------------------
Function loadPolizaCaucion(pIdPedido, ByRef saldoPDCPendiente, ByRef saldoPDCCargado, ByRef saldoPDCTotal, ByRef monedaPDC, ByRef IdPdc)
	Dim strSQL, auxImporte, auxCargado, auxTotal
	strSQL = "SELECT IDPDC, IMPORTE, CDMONEDA, ESTADO FROM TBLPOLIZASCAUCION WHERE IDPEDIDO = "& pIdPedido & " AND ESTADO <> "& ESTADO_PDC_ANULADA
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    while not rs.EoF	
	   auxMoneda = rs("CDMONEDA")
	   auxImporte = Cdbl(rs("IMPORTE"))	   
	   'Muestro en la parte de PDC Pendientes todas aquellas que se encuentren en ese estado, solo puede tener 1 PDC en este estado
	   if(rs("ESTADO") = ESTADO_PDC_PENDIENTE)then
	        saldoPDCPendiente =  auxImporte
	        IdPdc = Cdbl(rs("IDPDC"))
       end if
	   if(rs("ESTADO") <> ESTADO_PDC_PENDIENTE)then
	   'En caso de que no sean Pendientes, asumo que estan cargadas (Recibida,Vencida,Devuelta)		  	
		  	auxCargado = auxCargado + auxImporte		  	
	   end if
	rs.MoveNext
	wend	
	monedaPDC = auxMoneda
	saldoPDCCargado = auxCargado
	saldoPDCTotal = saldoPDCCargado + saldoPDCPendiente    
End Function
'----------------------------------------------------------------------------------------------------------------
'Esta funcion se encarga de ver si el PCP necesita un AFE, sin importar si lleva Partida o no.
Function checkPCPrequiredAFE(pIdPedido,pImporte,pMoneda)
	Dim auxImporte, rtrn
	rtrn = false
	if pImporte > 0 then
		'Se debe obtener el importe ganador del Pedido (en dolares)
		auxImporte = pImporte
		if (pMoneda = MONEDA_PESO) then	auxImporte = round(pImporte / getTipoCambio(MONEDA_DOLAR, ""),0)
		if (necesitaAFE(0, pIdPedido, 0,auxImporte,0,0)) then rtrn = true
	end if	
	checkPCPrequiredAFE = rtrn
End Function				
'---------------------------------------------------------------------------------------------------------------
'Controla la Partida Presupuestaria de la Planilla Presupuestaria
Function controlPartidaPresupuestariaPCP(pIdObra, pIdArea, pIdDetalle)
	Dim strSQL, conn, rsBudget,ret
	ret = true
	'Primero verifico si la PCP ya viene con una Partida cargada
	if (pct_idObra = 0) then
		'En caso de que no traiga una Partida cargada, controlo la que pueda ingresar el usuario
		if (pIdObra <> 0) then
			if (not isInversion(pIdObra)) then
				if ((pIdArea = 0) or (pIdDetalle = 0)) then Call setError(PCP_AREA_DET_OBLIGATORIO)
			end if
		else
			Call setWarning(OBRA_NO_SELECCIONADA)			
		end if
		if hayError() then ret = false
	end if			
	controlPartidaPresupuestariaPCP = ret			
End Function
'*******************************************************
'******	COMIENZO DE LA PAGINA
'*******************************************************
idPedido = GF_PARAMETROS7("idPedido",0,6)
Call initHeaderDB(idPedido)
if (not checkControlPCT()) then response.redirect "comprasAccesoDenegado.asp"

myAccion = GF_PARAMETROS7("accion","",6)
nroLinea = getCantidadCotizaciones(idPedido)

esExpo = (pct_idDivision = getDivisionID(CODIGO_EXPORTACION))

flagGrabar = false
if myAccion = "" then 
	'Leer desde base	
	RedimVarialbles(nroLinea)
	strSQL="SELECT * from TBLPCPDETALLE where IDPEDIDO=" & idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	myImporteSeleccion = 0
	myMonedaSeleccion = MONEDA_DOLAR
	index = 1
	while not rs.eof
			ITproveedor(index) = rs("IDPROVEEDOR")
			ITcaracteristica(index) = rs("CARACTERISTICAS")
			ITimporte(index) = rs("IMPORTE")
			ITimporte(index) = CDbl(ITimporte(index))/100
			ITmoneda(index) = rs("CDMONEDA")
			ITcondPago(index) = rs("CONDPAGO")
			ITnroLinea(index) = rs("NROSOBRE")
			if clng(rs("FECENTREGA")) = 0 then
				ITfecEntrega(index) = ""
			else	
				ITfecEntrega(index) = right(rs("FECENTREGA"),2) & "/" & mid(rs("FECENTREGA"),5,2) & "/" & left(rs("FECENTREGA"),4)			
			end if	
			if(CLng(ITproveedor(index)) = CLng(pct_idProveedorElegido))then 
			    myIdProveedor = pct_idProveedorElegido
			    myImporteSeleccion = ITimporte(index)
			    myMonedaSeleccion = ITmoneda(index)
			end if
			index = index + 1	
		    rs.movenext
	wend		
	Call loadPolizaCaucion(idPedido, myMontoPoliza, saldoPDCCargado, saldoPDCTotal, monedaPDC, IdPdc)		
	isPendiente = false		
	tienePoliza = false
	if(myMontoPoliza > 0)then
		isPendiente = true
		tienePoliza = true
	end if
	strSQL="SELECT * from TBLPCPCABECERA where IDPEDIDO=" & idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then			
		myComentarios = rs("COMENTARIOS")
		myComentarios = replace(myComentarios,ENTER_SYMBOL, chr(10))
		myComentarios = replace(myComentarios,"*","'")		
		strSQL = "Select * from TBLPCPFIRMAS where IDPEDIDO=" & pct_idPedido & " order by SECUENCIA"
		Call executeQueryDb(DBSITE_SQL_INTRA, rsFirmas, "OPEN", strSQL)
		'Las firmas se leen en secuencia creciente y se van llenando las posiciones de firma en orden.
        'Esto es así debido a las diferentes variantes que existen. Después el sistema determina que firmas muestra y cuales son solo a modo informativo ya que son firmas predeterminadas.
		if (not rsFirmas.eof) then
		    if (CInt(rsFirmas("SECUENCIA")) = PCP_FIRMA_RESPONSABLE) then
			    responsableCd = rsFirmas("CDUSUARIO")
			    myResponsable = getUserDescription(responsableCd)
			    rsFirmas.MoveNext()
			end if
		end if
		if (not rsFirmas.eof) then
		    if (CInt(rsFirmas("SECUENCIA")) = PCP_FIRMA_MIEMBRO1) then
			    member1Cd = rsFirmas("CDUSUARIO")			
			    member1 = getUserDescription(member1Cd)
			    rsFirmas.MoveNext()
			end if
		end if		
		if (not rsFirmas.eof) then
		    if (CInt(rsFirmas("SECUENCIA")) = PCP_FIRMA_MIEMBRO2) then
			    member2Cd = rsFirmas("CDUSUARIO")			
			    member2 = getUserDescription(member2Cd)
			    rsFirmas.MoveNext()
			end if
		end if				
	end if	
else
	'Leer desde pagina	
	myIdProveedor = GF_PARAMETROS7("SELPROV",0,6)	
	myComentarios = GF_PARAMETROS7("COM","",6)
	responsableCd = GF_PARAMETROS7("responsableCd","",6)
	if (responsableCd <> "") then myResponsable = getUserDescription(responsableCd)
					
	pct_idObra = GF_PARAMETROS7("cmbObra",0,6)	
	pct_idArea = GF_PARAMETROS7("idBudgetArea",0,6)
	pct_idDetalle = GF_PARAMETROS7("idBudgetDetalle",0,6)
	tienePoliza = false
	Call loadPolizaCaucion(idPedido, myMontoPoliza, saldoPDCCargado, saldoPDCTotal, monedaPDC, IdPdc)	
	myMontoPoliza = 0
	if (GF_PARAMETROS7("myPoliza",0,6) <> 0) then 
	    tienePoliza = true 'tiene tildado el CHECK	
	    myMontoPoliza = GF_PARAMETROS7("hidMontoPoliza",0 ,6)	    
	end if
	saldoPDCTotal = saldoPDCCargado + myMontoPoliza
	IdPdc = GF_PARAMETROS7("IdPdc",0 ,6)	
	isPendiente = GF_PARAMETROS7("isPendiente", "",6)	
	member1Cd = GF_PARAMETROS7("member1Cd","",6)	
	if (member1Cd <> "") then member1 = getUserDescription(member1Cd)	
	member2Cd = GF_PARAMETROS7("member2Cd","",6)
	if (member2Cd <> "") then member2 = getUserDescription(member2Cd)
	
	RedimVarialbles(nroLinea)
	for index = 1 to nroLinea
		ITproveedor(index) = GF_PARAMETROS7("PRO_" & index,0,6)
		ITcaracteristica(index) = GF_PARAMETROS7("CAR_" & index,"",6)
		ITimporte(index) = GF_PARAMETROS7("PRE_" & index,"",6)
		ITimporte(index) = Replace(ITimporte(index),",",".")
		if ITimporte(index) = "" then ITimporte(index) = 0
		ITmoneda(index) = GF_PARAMETROS7("MON_" & index,"",6)
		ITcondPago(index) = GF_PARAMETROS7("CON_" & index,"",6)
		ITnroLinea(index) = GF_PARAMETROS7("NRO_" & index,0,6)
		ITfecEntrega(index) = GF_PARAMETROS7("issuedate_" & index, "",6)
		if (ITfecEntrega(index) = "0") then	ITfecEntrega(index) = ""
		if(ITproveedor(index) = myIdProveedor)then 	
		    g_IndexSelect = index
		     myIdProveedor = ITproveedor(index)
			 myImporteSeleccion = ITimporte(index)
			 myMonedaSeleccion = ITmoneda(index)
        end if			    
        
	next
	'Se controla la planilla
	controlOK =	controlarPlanilla(responsableCd, member1Cd, member2Cd, ITfecEntrega, ITimporte,ITmoneda, myIdProveedor, saldoPDCTotal, ITproveedor, pct_idDivision,pct_idObra,pct_idArea,pct_idDetalle)
	if ((myAccion = ACCION_GRABAR) and (controlOK)) then		
	    'Para saber que firmas grabar, se determina el tipo de planilla.			    
	    pcpType = getPCPAuthorizationType(myImporteSeleccion, myMonedaSeleccion)    
		Call adminPCPFirmas(idPedido, PCP_FIRMA_RESPONSABLE, responsableCd)
		Call adminPCPFirmas(idPedido, PCP_FIRMA_MIEMBRO1, member1Cd)
		Call adminPCPFirmas(idPedido, PCP_FIRMA_MIEMBRO2, member2Cd)
		'Se agregan las firmas especiales pre definidas. 		
		'Se implementó de esta manera para que siempre se pase por todas las firmas y se borren las que no correspndan si por alguna razon se grabaron en la base.
		'-- Gerente del Sector
		auxUser = ""
		if (esExpo) then auxUser = FIRMA_NO_USER
		Call adminPCPFirmas(idPedido, PCP_FIRMA_GTE_SECTOR, auxUser)		
		'-- Gerente de Puerto
		auxUser = ""
		if (not esExpo) then auxUser = FIRMA_NO_USER
		Call adminPCPFirmas(idPedido, PCP_FIRMA_GTE_PUERTO, auxUser)		
        '-- Gerente de Compras
        Call adminPCPFirmas(idPedido, PCP_FIRMA_GTE_COMPRAS, FIRMA_NO_USER)		
        '-- Coordinador de Puertos 
        if ((not esExpo) and (pcpType <> PCP_TYPE_PURCHASE_SMALL)) then Call adminPCPFirmas(idPedido, PCP_FIRMA_SUP_PUERTOS, auxUser)                        
        '-- Director
        auxUser = ""
		if (pcpType = PCP_TYPE_PURCHASE_LARGE) then auxUser = FIRMA_NO_USER
		Call adminPCPFirmas(idPedido, PCP_FIRMA_DIRECCION, auxUser)
	    'Se graba la cabecera
		call addPCPCabecera(idPedido, myComentarios, "")
		'Se actualiza la cabecera con el proveedor seleccionado y el sector del responsable de firma.
		'Para los puertos se deja en cero para que se rija por el rol de la firma. 
	    idSector=0
	    if (esExpo) then idSector=getUserSector(responsableCd)
		strSQL="Update TBLPCTCABECERA set IDPROVEEDOR=" & myIdProveedor & ", IDSECTOR=" & idSector & ", ESTADO=" & ESTADO_PCT_EN_ANALISIS &", "&_
			   " IDOBRA = "& pct_idObra &",IDAREA = "& pct_idArea &", IDDETALLE = "& pct_idDetalle &" where IDPEDIDO=" & idPedido
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		for index = 1 to nroLinea
			call addPCPItems(idPedido, ITnroLinea(index), ITproveedor(index), ITcaracteristica(index), ITimporte(index)*100, ITmoneda(index), ITcondPago(index), ITfecEntrega(index))		
			if(ITproveedor(index) = myIdProveedor)then 	auxMoneda = ITmoneda(index)
		next		
		if(tienePoliza)then 
			if(isPendiente)then
				call updateSaldoPolizaCaucion(IdPdc, myMontoPoliza, auxMoneda)	
			else
				call addPolizaCaucion(idPedido, TIPO_PDC_POR_ADELANTO, myMontoPoliza, auxMoneda, session("MmtoSistema"), session("Usuario"),ESTADO_PDC_PENDIENTE)
			end if
		else
		    strSQL = "Delete from TBLPOLIZASCAUCION where IDPEDIDO=" & idPedido & " and ESTADO=" & ESTADO_PDC_PENDIENTE
		    Call executeQueryDb(DBSITE_SQL_INTRA, rsPDC, "EXEC", strSQL)
		end if
		if (esProforma(myIdProveedor)) then Call notifySuppliersDept(myIdProveedor)
		flagGrabar = true
	end if 
end if
if (not isNull(pct_idProveedor) and (myIdProveedor=0)) then 
	myIdProveedor = pct_idProveedorElegido
end if	
Call initProveedoresDB()
if nroLinea <> 0 then nroLinea = 1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ANALISIS COMPARATIVO DE OFERTAS</title>
<link rel="stylesheet" type="text/css" href="css/ActiSAIntra-1.css">	
<link rel="stylesheet" type="text/css" href="css/main.css">
<link rel="stylesheet" type="text/css" href="css/toolbar.css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/controles.js">		</script>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
	var currentId;
	var ch = new channel;	
	function MostrarCalendario(p_objID, funcSel, id) {
		currentId = id;
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
	function SeleccionarCalEmision(cal, date) {
		var str= new String(date);
		var myObj;
        document.getElementById("issuedateTXT_" + currentId).value = str;
	    document.getElementById("issuedate_" + currentId).value = str;
		
		myObj = document.getElementById("DelDate_" + currentId);
		myObj.style.visibility = 'visible';
	    
		if (cal) cal.hide();
	}
	function CerrarCal(cal) {
		cal.hide();
	}
	function submitPage(pAccion){
		document.getElementById("accion").value = pAccion; 
		document.getElementById("frmSel").submit(); 
	}
	function setValueSel(pObj, pIdProv, pMoneda, pNroLinea){
		document.getElementById("SELPROV").value = pIdProv;
		var myMoneda = document.getElementById("MON_" + pNroLinea).value
		if(myMoneda == '<%=MONEDA_PESO%>'){
			document.getElementById("cdmonedaPDC").innerHTML = '<%=TIPO_MONEDA_PESO%>';
			document.getElementById("cdmonedaPDC_confirmado").innerHTML = '<%=TIPO_MONEDA_PESO%>';
			document.getElementById("cdmonedaPDC_total").innerHTML = '<%=TIPO_MONEDA_PESO%>';			
		}else{
			document.getElementById("cdmonedaPDC").innerHTML = '<%=TIPO_MONEDA_DOLAR%>';
			document.getElementById("cdmonedaPDC_confirmado").innerHTML = '<%=TIPO_MONEDA_DOLAR%>';
			document.getElementById("cdmonedaPDC_total").innerHTML = '<%=TIPO_MONEDA_DOLAR%>';
		}
		checkSignatures(pNroLinea);
	}
	function SetValueFecha(pThat, pLine){
		var myObj;
		myObj = document.getElementById("issuedate_" + pLine);
		myObj.value = 0;
        myObj = document.getElementById("issuedateTXT_" + pLine);
		myObj.value = 'A convenir';		
		pThat.style.visibility = 'hidden';
	}
	
	function bodyOnLoad(){				
			var tb = new Toolbar('toolbar');
			tb.addButton("toolbar-save","Guardar","submitPage('<% =ACCION_GRABAR %>')");
			tb.addButton("toolbar-control","Controlar","submitPage('<% =ACCION_CONTROLAR %>')");
			tb.addButton("toolbar-cancel","Cancelar","window.close()");
			tb.draw();
		 <%	if pct_idObra > 0 then %>
				cargarPartida('<%=pct_idObra%>','<%=pct_idArea%>','<%=pct_idDetalle%>');				
		 <%	end if %>
		 
		 loadSignatureTable('<% =myImporteSeleccion %>', '<% =myMonedaSeleccion %>');
		 
	}
		
	var old_mon = '<% =myMonedaSeleccion %>';
	var old_pre = '<% =myImporteSeleccion %>';
	
	function checkSignatures(pLinea) {
	    var sel = document.getElementById("SEL_" + pLinea).checked;
	    var pre = document.getElementById("PRE_" + pLinea).value;
	    var mon = document.getElementById("MON_" + pLinea).value;
	    
	    if (sel) {
	        if ((old_mon != mon) || (old_pre != pre)) {
	            //Actualizo la tabla de firmas para la nueva condicion.
	            old_pre = pre;
	            old_mon = mon;
	            loadSignatureTable(pre, mon);	            
	        }       
	    }
	}
	
	function loadSignatureTable(pImporte, pMoneda) {
	    document.getElementById("signatureTableDiv").innerHTML = "Loading...";
<%  if (esExpo) then      %>
	    ch.bind("comprasComparativoDeOfertasExpoSignAjax.asp?importe=" + pImporte + "&moneda=" + pMoneda, "signatureTableCallback()");
<%  else            %>	    
        ch.bind("comprasComparativoDeOfertasSignAjax.asp?importe=" + pImporte + "&moneda=" + pMoneda, "signatureTableCallback()");
<%  end if           %>        
		ch.send();
	}
	
	function signatureTableCallback(){
		document.getElementById("signatureTableDiv").innerHTML = ch.response();		
		var msRespTec = new MagicSearch("", "responsable", 40, 2, "comprasStreamElementos.asp?tipo=personas");
		msRespTec.setToken(";");
		msRespTec.onBlur = seleccionarRE;				
		msRespTec.setValue(document.getElementById("responsableDs").value);		
				
	    var msMember1 = new MagicSearch("", "member1", 40, 2, "comprasStreamElementos.asp?tipo=personas");
	    msMember1.setToken(";");
	    msMember1.onBlur = seleccionarM1;
	    msMember1.setValue(document.getElementById("member1Ds").value);		
		
		if (document.getElementById("member2")) {
		    var msMember2 = new MagicSearch("", "member2", 40, 2, "comprasStreamElementos.asp?tipo=personas");
		    msMember2.setToken(";");
		    msMember2.onBlur = seleccionarM2;
		    msMember2.setValue(document.getElementById("member2Ds").value);	
		}
		
	}	
	
	function seleccionarRE(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("responsableCd").value = arr[0];
			document.getElementById("responsableDs").value = arr[1];			
			ms.setValue(arr[1]);
		} else {
			if (desc == "") {
			    document.getElementById("responsableCd").value = "";
			    document.getElementById("responsableDs").value = "";
			}
		}						
	}		
	
	function seleccionarM1(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("member1Cd").value = arr[0];
			document.getElementById("member1Ds").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") {
			    document.getElementById("member1Cd").value = "";							
			    document.getElementById("member1Ds").value = "";
            }			    
		}		
	}		
	function seleccionarM2(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("member2Cd").value = arr[0];
			document.getElementById("member2Ds").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") {
			    document.getElementById("member2Cd").value = "";							
			    document.getElementById("member2Ds").value = "";
            }			    
		}		
	}		
	function controlarMonto(){
		if(document.getElementById("myPoliza").checked){
			$("#myMontoPoliza").attr('disabled', false);
		}
		else{
			$("#myMontoPoliza").attr('disabled', true);
		}
	}
	function asignarMonto(pInputHid, pInputText){
		var tipoCambio = document.getElementById(pInputText).value;
		tipoCambio = tipoCambio.replace(/,/,".");
		document.getElementById(pInputHid).value = tipoCambio * 100;
	}
	function cargarPartida(idObra,idArea,idDetalle){
        if (idObra != '<%=OBRA_GEID %>' ){
		    ch.bind("almacenObtenerBudget.asp?idObra=" + idObra + "&idBudgetArea=" + idArea + "&idBudgetDetalle=" + idDetalle + "&accion=<%=ACCION_PROCESAR%>", "actualizarBudgetsCallback()");
		    ch.send();
        }
        else
        {
            document.getElementById("secBudgetDiv").innerHTML = "";
        }
	}
	function actualizarBudgetsCallback(){
		document.getElementById("secBudgetDiv").innerHTML = ch.response();
	}
	function readBudgetArea(me){		
		document.getElementById('idBudgetArea').value=$("#idBudgetDetalle option:selected").attr("alt");
	}
</script>
<style>
  	label { padding: 0; marign: 0; display: block; }
  	textarea { width: 100%; border: 0px; padding: 0px; }
  	.the-fix { }
	.celda {
		border-radius:8px 8px 8px 8px;
	}
</style>
</head>
<body onLoad="bodyOnLoad();">
<div id="toolbar"></div>
<form method="post" id="frmSel" action="comprasComparativoDeOfertas.asp?idPedido=<%=pct_idPedido%>">
    <div class="col66"></div>
    <br>
    <div class="tableaside size100"> <% call showErrors() %></div>
    <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="33%" align="center"> <% =GF_TRADUCIR("PUERTO") %></th>
                <th width="33%" align="center"> <% =GF_TRADUCIR("PEDIDO ") %></th>
                <th width="33%" align="center"> <% =GF_TRADUCIR("OBRA / TRABAJO") %> </th>
            </tr>
        </thead>
        <tbody>
    	    <tr>
                <td align="center"> <% =pct_dsDivision %> </td>
                <td align="center"> 
                    <% if isnull(pct_cdPedido) then
					        Response.write GF_TRADUCIR("Sin Pedido")
				       else
					        Response.write pct_tituloPedido & " (" & pct_cdPedido & ")"
				       end if  %>
                </td>
                <td align="center">
                <%  'Se permite la carga/actualizacion de la partida presupuestaria.
					Set rsObras = obtenerListaObras("", "", "",pct_idDivision, OBRA_ACTIVA) %>
					<select id="cmbObra" name="cmbObra">
						<option value="0" onclick="cargarPartida(0,0,0);">- <% =GF_TRADUCIR("Seleccione Partida") %> -
				<%		while (not rsObras.eof)	%>
							<option value="<% =rsObras("IDOBRA") %>" <% if (Cdbl(rsObras("IDOBRA")) = Cdbl(pct_idObra)) then response.write "selected='true'" %> onclick="cargarPartida(this.value,0,0);"><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %>
				<%			rsObras.MoveNext()			
						wend %>
                        <option value="<%=OBRA_GEID %>" <% if (OBRA_GEID = Cdbl(pct_idObra)) then response.write "selected='true'" %> onclick="cargarPartida(this.value,0,0);" ><% =OBRA_GECD %> - <% =GF_TRADUCIR(OBRA_GEDS) %>
					</select>
					<span id="secBudgetDiv"></span>
					<input type="hidden" id="idBudgetArea" name="idBudgetArea">
                </td>
            </tr>
        </tbody>
  </table>
  <br>
  <table class="datagrid" width="95%" align="center">
    <thead>
        <tr>
            <th align="center">NºSOBRE</th>
            <th align="center">PROVEEDOR</th>
            <th align="center">CARACTERISTICAS</th>
            <th align="center">PRECIOS</th>
            <th align="center">CONDICIONES DE PAGO</th>
            <th align="center">FECHA DE ENTREGA</th>
            <th align="center">OBSERVACIONES</th>
            <th align="center">SEL</th>
        </tr>
    </thead>
	<tbody>
    	<tr>
        <%  while (readNextProveedorDB()) 		
		        myQuitar = " visibility:hidden; "
		        checked = ""
		        if cint(ITnroLinea(nroLinea)) = 0 then ITnroLinea(nroLinea) = nroLinea
		        if pct_pathCotizacion = ACCION_PCT_RETIRARSE then pct_hayCotizacion = false
		        if not pct_hayCotizacion then
			        myNoCotizaState = "Disabled"
			        myNoCotizaValue = ACCION_PCT_RETIRARSE
			        ITimporte(nroLinea) = 0
			        ITmoneda(nroLinea) = MONEDA_PESO
			        ITcondPago(nroLinea) = myNoCotizaValue
			        ITcaracteristica(nroLinea) = myNoCotizaValue
			        ITfecEntrega(nroLinea) = myNoCotizaValue
		        else
			        pct_Abierto = CotizacionIsOpen(pct_idPedido, pct_idProveedor)
			        if ITfecEntrega(nroLinea) = "" then
				        fechaLabel = "A convenir"
			        else
				        fechaLabel = ITfecEntrega(nroLinea)
				        myQuitar = " visibility:visible; "				
			        end if			
			        if (CLng(myIdProveedor) = CLng(pct_idProveedor)) then
				        checked = "checked"
				        auxMoneda = getSimboloMoneda(ITmoneda(nroLinea))				
			        end if
			        myNoCotizaState = ""
			        myNoCotizaValue = ""
		        end if %>
            <td align="center">
                <%=nroLinea%>
				<input type="hidden" id="NRO_<%=nroLinea%>" name="NRO_<%=nroLinea%>" value="<%=ITnroLinea(nroLinea)%>">
			</td>
			<td><%=pct_dsProveedor%>
				<input type="hidden" name="PRO_<%=nroLinea%>" value="<%=pct_idProveedor%>">
			</td>
			<td align="center"><input type="" name="CAR_<%=nroLinea%>" id="CAR_<%=nroLinea%>" maxLength="50" value="<%=ITcaracteristica(nroLinea)%>" <%=myNoCotizaState%>></td>
			<td align="center">
				<input type="text" id="PRE_<%=nroLinea%>" style="text-align:right;" align="right" name="PRE_<%=nroLinea%>" maxlength="15" size="10" value="<% =ITimporte(nroLinea) %>" onKeyPress="return controlIngreso(this, event, 'I')" onblur="checkSignatures('<%=nroLinea%>')" <%=myNoCotizaState%>>
				<select id="MON_<%=nroLinea%>" name="MON_<%=nroLinea%>" onchange="checkSignatures('<%=nroLinea%>')">
					<option value="<% =MONEDA_DOLAR %>" <% if(ITmoneda(nroLinea) = MONEDA_DOLAR) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_DOLAR) %>
					<option value="<% =MONEDA_PESO %>"  <% if(ITmoneda(nroLinea) = MONEDA_PESO) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_PESO) %>							
				</select>
			</td>
			<td align="center"><input type="text" maxlength="38" id="CON_<%=nroLinea%>" name="CON_<%=nroLinea%>" value="<%=ITcondPago(nroLinea)%>" <%=myNoCotizaState%>></td>
			<td align="left">
				<% if pct_hayCotizacion then %>
					<a id="LNK_<%=nroLinea%>" href="javascript:MostrarCalendario('imgEmision', SeleccionarCalEmision, <%=nroLinea%>)">
						<img id="imgEmision" src="images/calendar-16.png">
					</a>
				    <img id="DelDate_<%=nroLinea%>" style='cursor:pointer;<%=myQuitar%>' src='images/compras/cancel-16x16.png' title='Quitar Fecha' onClick="SetValueFecha(this, <%=nroLinea%>)">
                    &nbsp; &nbsp;<input id="issuedateTXT_<%=nroLinea%>" name="issuedateTXT_<%=nroLinea%>" type="" value="<%=fechaLabel%>" size="10" readonly=readonly />				    
                <% end if %>
                <input type="hidden" id="issuedate_<%=nroLinea%>" name="issuedate_<%=nroLinea%>" value="<%=ITfecEntrega(nroLinea)%>">
			</td>
			<td>
			    <% 'Verifico si alguna de las coptizacion del proveedor fue presentada fuera del palzo.
			        Set rsCotizaciones = getCotizaciones(pct_idPedido, pct_idProveedor)
			        blnFuera = false
			        while ((not rsCotizaciones.eof) and (not blnFuera))
			            if  (GF_DTEDIFF(rsCotizaciones("FECHAPRESENTACION"), GF_DTE2FN(pct_FechaCierre), "D") < 0) then 
			    %>
			            <label class="reg_header_error round_border_all" title="Cotizacion cargada fuera de termino" style="cursor:pointer; padding:2px;">
			                Cotizacion cargada fuera de termino
			            </label>                			            
			    <%          blnFuera = true
			            end if 
			            rsCotizaciones.MoveNExt()
			        wend
			    %> 
			</td>
			<td align="center">
				<%	if pct_hayCotizacion and pct_Abierto then %>
						<input style='cursor:pointer;border:none;' title="Seleccionar Cotizacion" name="SELOPT" id="SEL_<%=nroLinea%>" type='radio' onClick="setValueSel(this, <% =pct_idProveedor %>,'<% =getSimboloMoneda(ITmoneda(nroLinea)) %>', '<%=nroLinea%>')" <%=checked%>>
				<%	end if %>
			</td>
		</tr>
    <%	nroLinea = nroLinea + 1
	wend  %>
    </tbody>
</table>
<br>
<table class="datagrid" width="95%" align="center">
    <thead>
        <tr class="alertmsj">
            <td colspan="3" class="reg_Header_Warning"> 
                <img src="images/compras/action_warning-16x16.png" alt="warning-16" />
                &nbsp; <% =GF_TRADUCIR("Una vez que la planilla Comparativa se encuentra aprobada, no se podr&aacuten agregar Polizas de cauci&oacuten") %>. 
            </td>
        </tr>
        <tr>
            <td colspan="3" align="center"> P&OacuteLIZA DE CAUCI&OacuteN (PDC) </td>
        </tr>
    </thead>
    <tbody>
    	<tr>
            <td width="33%" align="left" valign="middle">
            <%  if(isPendiente)then %>
				<%= GF_TRADUCIR("Saldo Pendiente: " ) %>
			<%  else  %>
			    <%= GF_TRADUCIR("Nueva Poliza: " ) %>
			<%  end if %>
                <input style="border:none;cursor:pointer;" type="checkbox" onclick="controlarMonto()" name="myPoliza" id="myPoliza" value="1" <%if (tienePoliza) then Response.Write "Checked"%> >
                <span id="cdmonedaPDC" ><%=auxMoneda%></span>
                <input type="text" id="myMontoPoliza" name="myMontoPoliza" value="<% =myMontoPoliza/100 %>" <% if (not tienePoliza) then Response.Write " DISABLED " %>  style="text-align:right;" onKeyPress="return controlIngreso(this, event, 'I')" onBlur="asignarMonto('hidMontoPoliza','myMontoPoliza');">
			    <input type="hidden" id="hidMontoPoliza" name="hidMontoPoliza" value="<% =myMontoPoliza %>">
            </td>
            <td width="33%" align="left">
                <%= GF_TRADUCIR("PDC Recibidas: ") %>
				<span id="cdmonedaPDC_confirmado" ><%=auxMoneda%></span>
				<%= saldoPDCCargado/100 %>
                <input type="hidden" name="hidMontoPolizaCargada" id="hidMontoPolizaCargada" value=<%=saldoPDCCargado%>>
            </td>
            <td width="33%" align="center"> 
                <%= GF_TRADUCIR("Importe Total de PDC: ") %>
				<span id="cdmonedaPDC_total" ><%=auxMoneda%></span>
				<%= saldoPDCTotal/100%>					
				<input type="hidden" name="hidMontoPolizaTotal" id="hidMontoPolizaTotal" value=<%=saldoPDCTotal%>>
            </td>
        </tr>
        <input type="hidden" name="isPendiente" id="isPendiente" value='<%=isPendiente%>' />
		<input type="hidden" name="monedaPDC" id="monedaPDC" value='<%=monedaPDC%>' />
		<input type="hidden" name="idPdc" id="idPdc" value='<%=idPdc%>' />
    </tbody>
</table>
<br><br>

<table width="95%" align="center" border="0">
    <tr>
        <td align="center">&nbsp;</td>
        <td width="45%" align="center">
            <table class="datagrid" width="95%" align="left">
   			    <thead>
        			<tr>
          				<th style="border-radius: 8px 8px 0 0" align="center"> COMENTARIOS/SUGERENCIAS T&EacuteCNICAS </th>
        			</tr>
    			</thead>
				<tbody>
    				<tr>
           				<td align="center"> <textarea name="COM" id="COM" type="text" rows="8" cols="150"><%=myComentarios%></textarea></td>
					</tr>
				</tbody>
			</table>
        </td>        
        <td width="33%" align="center" valign="top">
            <div id="signatureTableDiv"></div>
        </td>
    </tr>
</table>

<input type="hidden" name="accion" id="accion">		
<input type="hidden" name="SELPROV" id="SELPROV" value="<% =myIdProveedor %>">
<input type="hidden" id="responsableCd" name="responsableCd" value="<%=responsableCd%>">
<input type="hidden" name="responsableDs" id="responsableDs" value="<% =myResponsable %>">
<input type="hidden" id="member1Cd" name="member1Cd" value="<% =member1Cd %>">
<input type="hidden" id="member1Ds" name="member1Ds" value="<% =member1 %>">
<input type="hidden" id="member2Cd" name="member2Cd" value="<% =memeber2Cd %>">
<input type="hidden" id="member2Ds" name="member2Ds" value="<% =memeber2 %>">
</form>	
</body>
</html>
