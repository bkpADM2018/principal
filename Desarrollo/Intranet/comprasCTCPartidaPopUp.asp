<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosCTz.asp"-->

<%
'--------------------------------------------------------------------------------------------
Function controlarPartidaDiferente(pIdContrato,pIdobra,pFechcCierre,pFechaInicio)
    Dim strSQL,rtrn
	rtrn = true
    'CONTROLO QUE NO HALLA PARTIDAS DEL CONTRATO PARA ESE PERIODO (Solo se permite abrir detalles)
	strSQL = "SELECT * FROM TBLCTCPARTIDAS WHERE IDCONTRATO = " & pIdContrato &_
			 "  AND (FECHAINICIO <= " & pFechcCierre & ") AND (FECHACIERRE >= " & pFechaInicio & ")" &_
			 "  AND IDOBRA <> " & pIdobra
	'Response.Write strSQL				
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.Eof then rtrn = false
	controlarPartidaDiferente	= rtrn	
End Function
'--------------------------------------------------------------------------------------------
' Funci�n:	 
'			   controlarFechaCTCPartida
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   18/10/2013
' Objetivo:
'			   Controla que esten correctos los datos de la fecha de la Partida.
'			   Que el periodo entre ambas fechas sea correcto y que no halla una partida activa durante el per�odo	
'			   elegido ya que se pueden solapar.
' Parametros:  
'			   pIdContrato		[int]	ID Contrato
'			   fechaCierre		[int]	fecha Cierre
'			   fechaInicio		[int]	fecha Inicio
' Devuelve:    true: Periodo libre - false: Periodo Ocupado
'--------------------------------------------------------------------------------------------
Function controlarFechaCTCPartida(pIdContrato,pIdobra,pAreaObra,pDetalleObra,pFechcCierre,pFechaInicio)
	Dim strSQL,rtrn
	rtrn = true
	if pFechaInicio <> 0 and pFechcCierre <> 0 then	
		if (GF_CONTROL_PERIODO_2(GF_FN2DTE(pFechaInicio), GF_FN2DTE(pFechcCierre)) <> 0) then rtrn = false
	end if
	controlarFechaCTCPartida = rtrn
End Function
'--------------------------------------------------------------------------------------------
' Funci�n:	 
'			   controlarPartidaDuplicada
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   19/09/2013
' Objetivo:
'			   Controla que la Partida elegida para el Contrato no se halla asignado previamente 
' Parametros:  
'			   pIdContrato		[int]	ID Contrato
'			   idObra			[int]	ID Obra
'			   areaObra			[int]	ID Area
'			   detalleObra		[int]	ID Detalle
' Devuelve:    true: Partida OK - false: Duplicada
'--------------------------------------------------------------------------------------------
Function controlarPartidaDuplicada(pIdContrato,idobra,areaObra,detalleObra)
	Dim strSQL,rtrn
	rtrn =true
	if not isModificacion then 
		strSQL = "SELECT * FROM TBLCTCPARTIDAS WHERE IDCONTRATO="&pIdContrato&_
				 "   AND IDOBRA="&idobra&" AND IDAREA ="&areaObra&" AND IDDETALLE ="&detalleObra
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.Eof then rtrn= false
	end if	
	controlarPartidaDuplicada = rtrn
End Function
'------------------------------------------------------------------------------------------------
'Controla los datos de la Partida
Function controlarCTCPartida(pIdContrato,idobra,areaObra,detalleObra,pFechaCierre,pFechaEmision, pCdMoneda, pImporteAsignado, pImporteDisponible)	
	Dim rtrn, auxPesosCEC, auxDolaresCEC, auxImporteCEC 
	rtrn = false
	if idobra > 0 then	    
		if (controlarPartidaDuplicada(pIdContrato,idobra,areaObra,detalleObra)) then
		    if (controlarPartidaDiferente(pIdContrato,idobra,pFechaCierre,pFechaEmision)) then
			    if(controlarFechaCTCPartida(pIdContrato,idobra,areaObra,detalleObra,pFechaCierre,pFechaEmision)) then			        
			        if (CDbl(pImporteAsignado) <= CDbl(pImporteDisponible)) then
			            Call readCTCTotalPagado(pIdContrato, idobra, areaObra, detalleObra, auxPesosCEC, auxDolaresCEC, False)
			            auxImporteCEC = auxPesosCEC
			            if (pCdMoneda = MONEDA_DOLAR) then auxImporteCEC = auxDolaresCEC			        
			            if (round(CDbl(pImporteAsignado)*100, 0) >= round(Cdbl(auxImporteCEC), 0)) then			        
				            rtrn = true
                        else
                            Call setError(CTZ_AJU_TOTAL_BAJO)
                        end if				        
				    else
				        Call setError(CTC_SALDO_INSUFICIENTE)
				    end if
			    else
				    Call setError(PERIODO_ERRONEO)
			    end if
            else
                Call setError(CTC_PARTIDA_NO_UNICA)
            end if			    
		else
			Call setError(CTC_PARTIDA_YA_EXISTE)
		end if	
	else
		Call setError(OBRA_NO_EXISTE)
	end if
	controlarCTCPartida = rtrn
End Function
'--------------------------------------------------------------------------------------------------------------------
'Verifica si el usuario tiene permiso para agrgar una nueva Partida al Contrato
Function puedeAgregarPartidaCTC(pUser)
	Dim ret
	ret = false	
	if (getRolFirma(pUser, SEC_SYS_COMPRAS) = FIRMA_ROL_GTE_COMPRAS) then ret = true
	puedeAgregarPartidaCTC = ret
End Function
'--------------------------------------------------------------------------------------------------------------------

Dim idObra,idContrato,rs,accion,fechaPartida,hidFechaCierre,idDetalle,idArea,constolOK,flagGuardar, puedeAgregar 
DIm myMonedaCTC, myImporteCTC, importeAsignado, importeDisponible, importeAsignadoOriginal

Call GP_CONFIGURARMOMENTOS

idContrato = GF_PARAMETROS7("idContrato",0,6)
accion     = GF_PARAMETROS7("accion","",6)
idArea	   = GF_PARAMETROS7("idBudgetArea",0,6)
idDetalle  = GF_PARAMETROS7("idBudgetDetalle",0,6)
idObra     = GF_PARAMETROS7("idObra",0,6)
fecha	   = GF_PARAMETROS7("issuedate","",6)
importeAsignado	   = GF_PARAMETROS7("importeAsignado", 2,6)
importeAsignadoOriginal  = GF_PARAMETROS7("importeAsignadoOriginal", 2,6)
importeDisponible  = GF_PARAMETROS7("importeDisponible", 2,6)
importeDisponible  = importeDisponible + importeAsignadoOriginal

if ((fecha = "")or(idObra = 0)) then fecha = Left(session("MmtoDato"), 8)
fechaEmision	   = GF_PARAMETROS7("issuedateEmision","",6)
if (fechaEmision = "") then fechaEmision = Left(session("MmtoDato"), 8)
fechaPartida = GF_PARAMETROS7("fechaCierre","",6)
if(idobra = 0)then fechaPartida = ""
isModificacion = false
isModificacion = GF_PARAMETROS7("isModificacion","",6)
Call comprasControlAccesoCM(RES_CC)
flagGuardar= false
Set rs = readCTC(idContrato)
myImporteCTC = 0
myDivisionCTC = 0
if (not rs.eof) then
    myMonedaCTC = rs("CDMONEDA")
    myImporteCTC = rs("IMPORTEPESOS")        
    if (myMonedaCTC = MONEDA_DOLAR) then myImporteCTC = rs("IMPORTEDOLARES")       
    myDivisionCTC = rs("IDDIVISION")
end if    
if isFormSubmit() then
	constolOK = controlarCTCPartida(idContrato,idObra,idArea,idDetalle,fecha, fechaEmision, myMonedaCTC, importeAsignado, importeDisponible)	
	if ((accion = ACCION_GRABAR) and (constolOK)) then
		flagGuardar = true
		if (CDbl(myImporteCTC) > 0) then
		    Call grabarPartidaCTC(idContrato, idObra, idArea, idDetalle, fecha, fechaEmision, myMonedaCTC, importeAsignado*100, session("Usuario"))            		    
        end if		    
	end if
end if
Set rs = readCTCPartida(idContrato)
puedeAgregar = puedeAgregarPartidaCTC(session("Usuario"))
%>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">	
	<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
	<link rel="stylesheet" href="css/main.css" type="text/css"> 
	
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>	
	<script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/formato.js"></script>
	<script type="text/javascript" src="scripts/controles.js"></script>
	<script type="text/javascript" src="scripts/date.js"></script>
	
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
	<script type="text/javascript">
	var myDate = formatDate(new Date(),"dd/MM/yyyy");
	var ch = new channel;	
	var puw;
	function bodyOnLoad(){		
		var	tb = new Toolbar('toolbar', 4, "images/compras/");
		<%if (puedeAgregar) then %>
		idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
		idBtnControl = tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");
		<% end if %>
		tb.addButton("Close-16x16.png", "Cerrar", "cerrar()");
		tb.draw();
		<% if (idObra > 0) then %>
			actualizarBudgets(<%=idObra%>, <%=idArea%>, <%=idDetalle%>);		
			<% if isModificacion then %>
				document.getElementById("rowFechaInicio").style.display = "none";
				document.getElementById("divStrObra").style.display = "block";
			<% 	if (idObra <> OBRA_GEID) then
			        Call loadDatosObra(idObra, auxCdObra, auxDSObra, "", "", "", "", "", "", "", "", "", "")    			   
			    else
			        auxCdObra = OBRA_GECD
			        auxDSObra = OBRA_GEDS
			    end if
			%>
				document.getElementById("divStrObra").innerHTML = "<% =auxCdObra + " - " + auxDSObra %>";
				document.getElementById("cmbObra").style.display = "none";
				document.getElementById("secBudgetDiv").style.display = "none";
			<% end if %>
		<%end if%>
		<% if(flagGuardar)then %>
			document.getElementById("msgGuardado").innerHTML = "";
			document.getElementById("msgGuardado").innerHTML="<% =GF_TRADUCIR("Se guardo correctamente") %>";
			document.getElementById("msgGuardado").className = "TDSUCCESS";	
		<% end if %>	
	}
	
	function cerrar() {
		puw = getObjPopUp("popUpModificarObra");
		puw.hide();
	}
	
	function actualizarBudgetsCallback(){	    
        document.getElementById("secBudgetDiv").innerHTML = ch.response();     
	}
	
	/*function cargarPartida: carga los datos de una Partida(fecha cierre y area/detalle) 
					  idObra: Id Obra	*/
	function cargarPartida(idObra){
		if(idObra > 0){
			actualizarFechaCierre();
			actualizarBudgets(idObra, 0, 0);
		}
	}	
		
	/*function actualizarBudgets: se encarga de actualizar los datos de la Obra, llama por Ajax al Area y Detalle
								  de la Obra seleccionada
						  idObra: Obra seleccionada
						  idArea: Area seleccionada
					   idDetalle: Detalle seleccionado			*/
	function actualizarBudgets(idObra, idArea, idDetalle){
	    ch.bind("almacenObtenerBudget.asp?idObra=" + idObra + "&idBudgetArea=" + idArea + "&idBudgetDetalle=" + idDetalle + "&accion=<%=ACCION_PROCESAR%>", "actualizarBudgetsCallback()");
	    ch.send();        
	}
	/*function actualizarFechaCierre: se encarga de actualizar la fecha de Cierre de la obra seleccionada */
	function actualizarFechaCierre(){
		var myfecha = $("#cmbObra option:selected").attr("alt");		
		document.getElementById("fechaCierre").value = myfecha
		document.getElementById("issuedateDiv").innerHTML = myfecha;
		var fecha = myfecha.substr(6,4) + myfecha.substr(3,2) + myfecha.substr(0,2)
		document.getElementById("issuedate").value = fecha;						
		document.getElementById("issuedateDivEmision").innerHTML = myDate;
		document.getElementById("issuedateEmision").value = myDate.substr(6,4) + myDate.substr(3,2) + myDate.substr(0,2);						
	}
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("frm").submit();
	}
	function MostrarCalendario(p_objID, funcSel){
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
		
	function CerrarCal(cal){		
		cal.hide();
	}

	function SeleccionarCalCierre(cal, date) {		
		var flag = false;
		var str= new String(date);
		fechaCierre = document.getElementById("fechaCierre").value;
		if(fechaCierre != ''){
			//Valido que la fecha seleccionada no sea mayor a la del Cierre de la Partida
			var rtrn = compareDates(str,"dd/MM/yyyy", fechaCierre,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha de Vencimiento no puede ser mayor a la fecha de Cierre de la Obra!");
				str = fechaCierre;
				flag = true;	
			}
		}	
		if(flag == false){
			//Valido que la fecha seleccionada no sea menor a la fecha Actual
			var rtrn = compareDates(myDate,"dd/MM/yyyy", str,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha seleccionada no puede ser anterior a la actual!");
				str = myDate;
				if(fechaCierre != '') str = fechaCierre;
			}			
		}
		document.getElementById("issuedateDiv").innerHTML = str;
		document.getElementById("issuedate").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();
	}

	function SeleccionarCalEmision(cal, date) {		
		var flag = false;
		var str= new String(date);
		fechaCierre = document.getElementById("fechaCierre").value;
		if(fechaCierre != ''){
			//Valido que la fecha seleccionada no sea mayor a la del Cierre de la Partida
			var rtrn = compareDates(str,"dd/MM/yyyy", fechaCierre,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha de Emision no puede ser mayor a la fecha de Cierre de la Obra!");
				str = myDate;
				flag = true;	
			}
		}	
		if(flag == false){
			//Valido que la fecha seleccionada no sea menor a la fecha Actual
			var rtrn = compareDates(myDate,"dd/MM/yyyy", str,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha seleccionada no puede ser anterior a la actual!");
				str = myDate;				
			}			
		}
		document.getElementById("issuedateDivEmision").innerHTML = str;
		document.getElementById("issuedateEmision").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();
	}
	
	/* function readBudgetArea(me) : Esta funcion asigna a un elemento Hidden el valor del Area selccionada, se dispara cuando
								el combo box del Area/Detalle pierde el focus */
	function readBudgetArea(me){		
		document.getElementById('idBudgetArea').value=$("#idBudgetDetalle option:selected").attr("alt");
	}
	/* function modificarPartidaCTC : carga los valores de la Partida seleccionada al formulario para poder editarlos	*/
	function modificarPartidaCTC(pIdObra, pCdObra, pDsObra, pArea, pDetalle, pFechaVto, pFechaInicio, pImporte){
		document.getElementById("isModificacion").value = true;
		if (pFechaInicio != 0){
			document.getElementById("issuedateDivEmision").innerHTML = pFechaInicio;
			document.getElementById("issuedateEmision").value = pFechaInicio.substr(6,4) + pFechaInicio.substr(3,2) + pFechaInicio.substr(0,2);
		}
		else{			
			document.getElementById("issuedateEmision").value = 0;
		}	
		document.getElementById("rowFechaInicio").style.display = "none";
		document.getElementById("issuedateDiv").innerHTML = pFechaVto;
		document.getElementById("issuedate").value = pFechaVto.substr(6,4) + pFechaVto.substr(3,2) + pFechaVto.substr(0,2);
		document.getElementById("divStrObra").innerHTML = pCdObra + " - " + pDsObra
		document.getElementById("divStrObra").style.display = "block";		
		actualizarBudgets(pIdObra, pArea, pDetalle);
		document.getElementById("cmbObra").style.display = "none";
		document.getElementById("idObra").value = pIdObra;
		document.getElementById("secBudgetDiv").style.display = "none";
		document.getElementById("fechaCierre").value = $("#cmbObra option:selected").attr("alt");
		document.getElementById("importeAsignado").value = pImporte/100;
		document.getElementById("importeAsignadoOriginal").value = pImporte/100;
	}	
	function eliminarPartidaCTC(pContrato, pObra, pArea, pDetalle){
		if (confirm("Esta seguro que desea eliminar esta Partida Presupuestaria?")) {
			ch.bind("comprasCTCPartidaAjax.asp?idContrato="+pContrato+"&idObra="+pObra+"&idArea="+pArea+"&idDetalle="+pDetalle+"&accion=<%=ACCION_BORRAR%>", "eliminarBudgetsCallback()");
			ch.send();
		}	
	}
	function eliminarBudgetsCallback(){
		submitInfo();
	}
	function eligeObra() {
	    document.getElementById("idObra").value = $("#cmbObra option:selected").val();	    
	}
</script>
</head>
<BODY onLoad="bodyOnLoad()">	
	<div id="toolbar"></div>
	<br>
	<form name="frm" id="frm"> 			
	<%if (puedeAgregar) then
		call showErrors() %>
		<table id="tableBudget" width="100%" border="0" align="center" class="reg_header">
		<tr>
			<td colspan="3">
				<div id="msgGuardado" align="center" class="TDBAJAS"></div>
			</td>						
		</tr>
		<tr>	
			<td class="reg_header_nav" colspan="3"><% =GF_TRADUCIR("Nuevo") %></td>
		</tr>		
		<tr>			
			<td class="reg_header_navdos" width="20%"><% =GF_TRADUCIR("Ptda. Presup.:") %></td>
			<td colspan="2">				
			<%	Set rsObras = obtenerListaObras("", "", "", myDivisionCTC, OBRA_ACTIVA)	%>
				<select id="cmbObra" name="cmbObra" onchange="javascript:eligeObra()">
					<option value="0" onclick="cargarPartida(0);">- <% =GF_TRADUCIR("Seleccione Partida") %> -
			<%		while (not rsObras.eof)
						hidFechaCierre = Cdbl(rsObras("FECHAFIN"))
						if (Cdbl(rsObras("FECHAAJUSTADA")) <> 0) then hidFechaCierre = Cdbl(rsObras("FECHAAJUSTADA"))	%>
						<option value="<% =rsObras("IDOBRA") %>" alt="<%=GF_FN2DTE(hidFechaCierre)%>" <% if (rsObras("IDOBRA") = idObra) then response.write "selected='true'" %> onclick="cargarPartida(this.value);"><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %>
			<%			rsObras.MoveNext()			
					wend 			    
			%>
					<option value="<% =OBRA_GEID %>"  <% if (OBRA_GEID = Cdbl(idObra)) then response.write "selected='true'" %>><% =OBRA_GECD %> - <% =GF_TRADUCIR(OBRA_GEDS) %>
				</select>
				<span id="secBudgetDiv"></span>
				<input type="hidden" id="idObra" name="idObra" value="<% =idObra %>">				
				<input type="hidden" id="idBudgetArea" name="idBudgetArea">				
				<input type="hidden" id="fechaCierre" name="fechaCierre" value="<%=fechaPartida%>">					
				<div id="divStrObra" name="divStrObra" style="display:none;">
				<%  if isModificacion then 
						Set rsObra = obtenerListaObras(idObra, "", "","","") 
						if (not rsObra.eof) then response.write rsObra("CDOBRA") & " - " & rsObra("DSOBRA")						
					end if	 %>
				 <div>	
			</td>
		</tr>		
		<tr id="rowFechaInicio">
			<td class="reg_header_navdos"><% =GF_TRADUCIR("Fecha Inicio:") %></td>
			<td>
				<a href="javascript:MostrarCalendario('imgLimiteInicio', SeleccionarCalEmision)"><img id="imgLimiteInicio" src="images/DATE.gif"></a>
			</td>
			<td>
				<div id="issuedateDivEmision" class="labelStyle"><% =GF_FN2DTE(fechaEmision) %></div>
				<input type="hidden" id="issuedateEmision" name="issuedateEmision" value="<% =fechaEmision %>" />				
			</td>			
		</tr>
		<tr>
			<td class="reg_header_navdos"><% =GF_TRADUCIR("Fecha Cierre:") %></td>
			<td>
				<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalCierre)"><img id="imgLimite" src="images/DATE.gif"></a>
			</td>
			<td>
				<div id="issuedateDiv" class="labelStyle"><% =GF_FN2DTE(fecha) %></div>
				<input type="hidden" id="issuedate" name="issuedate" value="<% =fecha %>" />				
			</td>			
		</tr>
		<tr>
			<td class="reg_header_navdos"><% =GF_TRADUCIR("Importe Asignado:") %></td>
			<td colspan="2">
				<% =getSimboloMoneda(myMonedaCTC) %><input style="text-align:right;" type="text" size="10" id="importeAsignado" name="importeAsignado" value="<% =importeAsignado %>" onkeypress="return controlIngreso(this, event, 'N')"/>
			</td>			
		</tr>
	</table>
	<br>
  <%end if %>
	<table id="tableBudget" width="100%" border="0" cellpadding="1" cellspacing="2" align="center" class="reg_header">				
		<tr>			
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Obra") %></td>
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Area-Detalle") %></td>			
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Fecha Inicio") %></td>
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Fecha Cierre") %></td>
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Importe Asignado") %></td>
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Saldo") %></td>
			<td class="reg_header_nav" align="center"><%= GF_TRADUCIR("Usuario") %></td>
			<td class="reg_header_nav" align="center">.</td>
			<td class="reg_header_nav" align="center">.</td>
		</tr>
		  <% myImporteAcumulado = 0
		    while not rs.EoF %>				
					<tr class=<%if ((CLng(rs("FECHAINICIO")) <= CLng(left(session("MmtoDato"), 8))) and (CLng(rs("FECHACIERRE")) >= CLng(left(session("MmtoDato"), 8)))) then Response.Write "reg_header_green" else Response.Write "reg_Header_navdos" end if%>>
						<td align="left">  <% =rs("CDOBRA") & " - " & rs("DSOBRA") %></td>
						<td align="center"><% if ((Cdbl(rs("IDAREA")) <> 0)and(Cdbl(rs("IDDETALLE")) <> 0)) then Response.Write Cdbl(rs("IDAREA")) &" - "& Cdbl(rs("IDDETALLE")) %></td>
						<td align="center"><% if (Cdbl(rs("FECHAINICIO")) <> 0) then Response.Write GF_FN2DTE(Cdbl(rs("FECHAINICIO"))) %></td>
						<td align="center"><% =GF_FN2DTE(Cdbl(rs("FECHACIERRE"))) %></td>				
						<td align="right"><% =getSimboloMoneda(rs("CDMONEDA")) & " " & GF_EDIT_DECIMALS(rs("IMPORTEASIGNADO"), 2) %></td>
						<td align="right"><% =getSimboloMoneda(rs("CDMONEDA")) & " " & GF_EDIT_DECIMALS(CDbl(rs("IMPORTEASIGNADO")) - CDbl(rs("IMPORTEGASTADO")), 2) %></td>
						<td align="center"><% =rs("CDUSUARIO") %></td>						
						    <td align="center"><img src="images/compras/edit-16x16.png" style="cursor: pointer" title="<% =GF_TRADUCIR("Editar") %>" onClick="modificarPartidaCTC('<%=rs("IDOBRA")%>','<%=rs("CDOBRA")%>','<%=rs("DSOBRA")%>','<%=Cdbl(rs("IDAREA"))%>','<%=Cdbl(rs("IDDETALLE"))%>','<%=GF_FN2DTE(Cdbl(rs("FECHACIERRE")))%>','<%=GF_FN2DTE(Cdbl(rs("FECHAINICIO")))%>', '<% =rs("IMPORTEASIGNADO") %>');"></td>
                        <%  Call readCTCTotalPagado(idContrato, rs("IDOBRA"), rs("IDAREA"), rs("IDDETALLE"), auxPesosCEC, auxDolaresCEC, False)			      
						    if (auxPesosCEC = 0) then %>						    
						    <td align="center"><img src="images/compras/CTZ_cancel-16x16.png" style="cursor: pointer" title="<% =GF_TRADUCIR("Eliminar") %>" onClick="eliminarPartidaCTC('<%=idContrato%>','<%=Cdbl(rs("IDOBRA"))%>','<%=Cdbl(rs("IDAREA"))%>','<%=Cdbl(rs("IDDETALLE"))%>');"></td>
                        <% end if %>						    
					</tr>
		<%		myImporteAcumulado = myImporteAcumulado + CDbl(rs("IMPORTEASIGNADO"))
		        rs.MoveNext()
			wend 
        %>		
            <tr class="rtotal">				
				<td colspan="4" align="right"><%= GF_TRADUCIR("TOTAL: ") %></td>		
				<td align="right"><% =getSimboloMoneda(myMonedaCTC) & " " &  GF_EDIT_DECIMALS(myImporteAcumulado, 2) %></td>		
			</tr>
			<tr>				
				<td colspan="8" align="center">&nbsp;</td>				
			</tr>	
			<tr class="rtotal">				
				<td colspan="4" align="right"><%= GF_TRADUCIR("IMPORTE NO ASIGNADO: ") %></td>				
				<% importeDisponible = CDbl(myImporteCTC) - CDbl(myImporteAcumulado) %>
				<td align="right"><% =getSimboloMoneda(myMonedaCTC) & " " &  GF_EDIT_DECIMALS(importeDisponible, 2) %></td>		
			</tr>	
		</table>
		<input type="hidden" id="accion" name="accion">
		<input type="hidden" id="idContrato" name="idContrato" value="<%=idContrato%>">
		<input type="hidden" id="fechaEmision" name="fechaEmision" value="<%=fechaEmision%>">	
		<input type="hidden" id="isModificacion" name="isModificacion" value="<%=isModificacion%>">
		<input type="hidden" id="importeDisponible" name="importeDisponible" value="<%=CDbl(importeDisponible)/100 %>">
		<input type="hidden" id="importeAsignadoOriginal" name="importeAsignadoOriginal" value="<% =importeAsignadoOriginal %>">
	</form>	
</body>
</html>