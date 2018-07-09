<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<%
'******************************************************************************************
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
'********************************************************************
'					INICIO PAGINA
'********************************************************************
dim pto,rsGeneral,accion,strSQLPro,flagReport, fdesde

totalVagones = 0
totalKilosNetos = 0
Call GP_CONFIGURARMOMENTOS()

g_strPuerto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", g_strPuerto, params)
accion = GF_PARAMETROS7("accion", "", 6)
fdesde = GF_PARAMETROS7("fdesde", "", 6)
if (fdesde = "") then fdesde = GF_FN2DTCONTABLE(session("MmtoSistema"))
call addParam("cdProducto", myCdProducto, params)
if not hayError() then
	Set rsGeneral = loadCTGNoEnviados(fdesde)
end if
'-------------------------------------------------------------------------------------------------------
function loadCTGNoEnviados(fecIni)
dim strSQL, rs
		 
strSQL = "SELECT T1.ETAPA, T1.CTG, T1.NUCARTAPORTE, T1.IDCAMION, (YEAR(T1.DTCONTABLE)*10000 + Month(T1.DTCONTABLE)*100 + DAY(T1.DTCONTABLE))  DTCONTABLE, " &_
         "  T1.CDPRODUCTO, T1.CDESTADO, T1.ESTADOARRIBO, T1.ESTADOCONFIRMACION, T1.ESTADORECHAZO, T1.ESTADODESVIO, "&_
         " MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, T1.HCDCTG, " &_
         "   ((SELECT VLPESADA FROM HPESADASCAMION WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION AND CDPESADA=1 AND SQPESADA=(SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION AND CDPESADA=1)) - " & _
         "   (SELECT VLPESADA FROM HPESADASCAMION WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION AND CDPESADA=2 AND SQPESADA=(SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION AND CDPESADA=2)) - " & _
         "   (SELECT VLMERMAKILOS FROM HMERMASCAMIONES WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION AND SQPESADA=(SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=T1.DTCONTABLE AND IDCAMION=T1.IDCAMION))) KILOSNETOS " & _
         " FROM ( " 
'        "	        SELECT '1' AS ETAPA, WS.CTG, WS.NUCARTAPORTE, HC.IDCAMION, WS.DTCONTABLE, HC.NUAUTSALIDA, HC.CDPRODUCTO, HCD.CDCOSECHA, "&_
'        "                  HC.CDESTADO, ESTADOARRIBO, ESTADOCONFIRMACION, ESTADORECHAZO, ESTADODESVIO, MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, HCD.CTG HCDCTG " & _
'		 "	        FROM HCAMIONESDESCARGA HCD " & _
'		 "	                INNER JOIN dbo.HCAMIONES HC" & _
'		 "	                        ON HC.CDESTADO IN (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_RECHAZADO & ") AND HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _		 
'		 "	                INNER JOIN  dbo.WSCTG_CAMIONES WS " & _
'		 "	                        ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG<>CAST(HCD.CTG AS INT)" & _
'        "	                WHERE HCD.DTCONTABLE='" & fecIni & "' "          
'strSQL = strSQL & "	UNION" & _                
strSQL = strSQL & " SELECT '2' AS ETAPA, HCD.CTG, HCD.NUCARTAPORTE, HC.IDCAMION, HCD.DTCONTABLE, HC.CDPRODUCTO, "&_
         "                  HC.CDESTADO, ESTADOARRIBO, ESTADOCONFIRMACION, ESTADORECHAZO, ESTADODESVIO, MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, HCD.CTG HCDCTG " & _
		 "	        FROM (Select * from HCAMIONESDESCARGA where DTCONTABLE='" & fecIni & "') HCD " & _
		 "	                INNER JOIN dbo.HCAMIONES HC" & _
		 "	                        ON HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _		 
		 "	                LEFT JOIN dbo.WSCTG_CAMIONES WS" & _
		 "	                        ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG=CAST(HCD.CTG AS INT)" & _
		 "                  WHERE ( (HC.CDESTADO IN (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ") AND (ESTADOCONFIRMACION is Null))" &_
		 "                           OR    (HC.CDESTADO IN (" & CAMIONES_ESTADO_RECHAZADO & ") AND (ESTADORECHAZO is Null)) )"
strSQL = strSQL & "	UNION" & _
		 "	        SELECT '3' AS ETAPA, HCD.CTG, HCD.NUCARTAPORTE, HC.IDCAMION, HCD.DTCONTABLE, HC.CDPRODUCTO, "&_
         "                  HC.CDESTADO, ESTADOARRIBO, ESTADOCONFIRMACION, ESTADORECHAZO, ESTADODESVIO, MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, HCD.CTG HCDCTG " & _
		 "	        FROM (Select * from HCAMIONESDESCARGA where DTCONTABLE='" & fecIni & "') HCD " & _
		 "	                INNER JOIN dbo.HCAMIONES HC" & _
		 "	                        ON HC.CDESTADO IN (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ") AND HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _		 
		 "	                INNER JOIN  (Select * from WSCTG_CAMIONES where ESTADOCONFIRMACION = " & WSCTG_PENDIENTE & ") WS" & _
		 "	                        ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG=CAST(HCD.CTG AS INT) "
strSQL = strSQL & "	UNION" & _
		 "	        SELECT '4' AS ETAPA, HCD.CTG, HCD.NUCARTAPORTE, HC.IDCAMION, HCD.DTCONTABLE, HC.CDPRODUCTO, "&_
         "                  HC.CDESTADO, ESTADOARRIBO, ESTADOCONFIRMACION, ESTADORECHAZO, ESTADODESVIO, MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, HCD.CTG HCDCTG " & _
		 "	        FROM (Select * from HCAMIONESDESCARGA where DTCONTABLE='" & fecIni & "') HCD " & _
		 "	                INNER JOIN dbo.HCAMIONES HC" & _
		 "	                        ON HC.CDESTADO IN (" & CAMIONES_ESTADO_RECHAZADO & ") AND HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _		 
		 "	                INNER JOIN  (Select *  from WSCTG_CAMIONES where ESTADORECHAZO = " & WSCTG_PENDIENTE & ") WS " & _
		 "	                        ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG=CAST(HCD.CTG AS INT) "
strSQL = strSQL & "	UNION" & _
		 "	        SELECT '5' AS ETAPA, HCD.CTG, HCD.NUCARTAPORTE, HC.IDCAMION, HCD.DTCONTABLE, HC.CDPRODUCTO, "&_
         "                  HC.CDESTADO, ESTADOARRIBO, ESTADOCONFIRMACION, ESTADORECHAZO, ESTADODESVIO, MMTOARRIBO, MMTOCONFIRMACION, MMTORECHAZO, MMTODESVIO, CDUSERARRIBO, CDUSERCONFIRMACION, CDUSERRECHAZO, CDUSERDESVIO, HCD.CTG HCDCTG " & _
		 "	        FROM (Select * from HCAMIONESDESCARGA where DTCONTABLE='" & fecIni & "') HCD " & _
		 "	                INNER JOIN dbo.HCAMIONES HC" & _
		 "	                        ON HC.CDESTADO IN (" & CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_RECHAZADO & ") AND HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _		 
		 "	                INNER JOIN  dbo.WSCTG_CAMIONES WS " & _
		 "	                        ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG=CAST(HCD.CTG AS INT)" & _
         "	                WHERE (   (ESTADORECHAZO > " & WSCTG_PENDIENTE & ")" &_
         "                             OR (ESTADOCONFIRMACION > " & WSCTG_PENDIENTE & ") )" 
 strSQL = strSQL & "	) T1 " &_
             "           INNER JOIN DEVICES_CODE DC " & _
             "                  ON T1.CDPRODUCTO=DC.CDINTERNO AND DC.CDDEVICE= " & DEVICE_CODE_AFIP
 strSQL = strSQL & " ORDER BY ETAPA, T1.NUCARTAPORTE ASC "
		 
		 
		 	
	'Response.Write strSQL
	'Response.End 
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)
	
	Set loadCTGNoEnviados = rs

end function
'------------------------------------------------------------------------------------------------
function getEstadoDS(pCdEstado)
dim rtrn
if isnull(pCdEstado) then 
	rtrn = "Null"
else
	if cint(pCdEstado) = 8 or cint(pCdEstado) = 6 then
		rtrn = "Descargado"
	elseif cint(pCdEstado) = 7 then
		rtrn = "<font color='red'>Rechazado</font>"
	else
		rtrn = "No Aplica"	
	end if	
end if
getEstadoDS = rtrn
end function
'------------------------------------------------------------------------------------------------
function getEstadoCTG(pCdEstado, pCdUsuario)
dim rtrn
if isNull(pCdEstado) then 
	rtrn = "Nulo"
else
	if cint(pCdEstado) = WSCTG_PENDIENTE then	
		rtrn = "Pendiente"
	elseif cint(pCdEstado) = WSCTG_CONFIRMADO then	
		rtrn = "Realizado Automatico"
	elseif cint(pCdEstado) = WSCTG_MANUAL then	
		rtrn = "Realizado Manual por " & pCdUsuario	
	elseif cint(pCdEstado) = WSCTG_EXENTO then	
		rtrn = "No Aplica"
	elseif cint(pCdEstado) = WSCTG_QUITADO then	
		rtrn = "Sacado de lista por " & pCdUsuario	
	else
		rtrn = "Inválido"
	end if	
end if
getEstadoCTG = rtrn
end function
'-------------------------------------------------------------------------------------------------------------
function getFecha(pFecha,pEtapa)
dim rtrn
if Cdbl(pEtapa) = "2" then
	rtrn = "No Aplica"
else
	rtrn = GF_FN2DTE(rtrn)
end if	
getFecha = rtrn
end function
'-------------------------------------------------------------------------------------------------------------
function printEncabezado(pEtapa)

    Dim myTitle, myClass
    
		Select case cint(pEtapa)
		    case 1
		        myTitle = GF_Traducir("Camiones con problemas que no se enviarán a la AFIP. (CTG Distintos par ala carta de porte en RM&D y la tabla WSCTG_CAMIONES)")		        
                myClass = "reg_header_error"		        
		    case 2
		        myTitle = GF_Traducir("Camiones con problemas que no se enviarán a la AFIP. (Estado WS Confirmado = Nulo y Estado WS Rechazo = Nulo )")		        
		        myClass = "reg_header_error"
		    case 3 
		        myTitle = GF_Traducir("Camiones Esperando su Confirmación Definitiva. (Estado WS Confirmado = 0)")		        
		        myClass = "reg_header_warning"
		    case 4
		    	myTitle = GF_Traducir("Camiones Esperando su Rechazo. (Estado WS Rechazo = 0)")
		    	myClass = "reg_header_warning"
		    case 5
		        myTitle = GF_Traducir("Camiones con Confirmación Definitiva / Rechazo completo. (Estado WS Confirmado > 0 o Estado WS Rechazo > 0)")		        
		End Select		
%>
<thead>
	<tr>
		<td colspan='12' class="<% =myClass %>"><% =myTitle %></td>
	</tr>
			<TR>
				<th align="center"><%=GF_Traducir("Fecha Contable")%></th>
				<th align="center"><%=GF_Traducir("Carta Porte")%></th>
				<th align="center"><%=GF_Traducir("CTG")%></th>
				<th align="center"><%=GF_Traducir("Id Camion")%></th>
				<th align="center"><%=GF_Traducir("Producto")%></th>
				<th align="center"><%=GF_Traducir("Estado RMD")%></th>
				<th align="center"><%=GF_Traducir("Kilos Netos")%></th>
				<th align="center"><%=GF_Traducir("Estado WS Arribo")%></th>
				<th align="center"><%=GF_Traducir("Estado WS Confirmado")%></th>
				<th align="center"><%=GF_Traducir("Estado WS Rechazado")%></th>
				<th align="center"><%=GF_Traducir("Estado WS Desvio")%></th>
				<th align="center"><%=GF_Traducir("Editar")%></th>
			</TR>
</thead>
<%
end function
%>

<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=9">

<title><%=GF_TRADUCIR("Puertos - CTGs No Informados a AFIP")%></title>
<link rel="stylesheet" href="../css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/main.css" type="text/css"> 

<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript">	
	var ch = new channel();		
	var changeFilters = false;
	var maxSegments;
	var currSegment=0;
	var optionReport;	
	var MS_X_DAY = 86400000 //Milisegundos por día.	
	var d = new Date();	
	
	
	function volver() {	
		location.href = "../puertosReportes.asp?pto=<%=g_strPuerto%>";
	}

	function habilitarLoading(pVisibility, pPosition){
		document.getElementById("imgLoading").style.position = pPosition;
		document.getElementById("imgLoading").style.visibility  = pVisibility;
		document.getElementById("lblLoading").style.position = pPosition;
		document.getElementById("lblLoading").style.visibility  = pVisibility;
		if (pVisibility=='visible')
			document.getElementById("actionLabel").style.visibility  = "hidden";
		else	
			document.getElementById("actionLabel").style.visibility  = "visible";
	}

	function submitInfo(){
		document.getElementById("frmSel").submit();
	}
	
		
	function CerrarCal(cal) {
		cal.hide();
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
	    calendar.setDateFormat("y-mm-dd");
	    calendar.showAtElement(elem);
	}
	
	function SeleccionarCalDesde(cal, date) {
		var str= new String(date);
		document.getElementById("fdesde").value = str;
		if (cal) cal.hide();
	}	
	function enabledEdit(pIndex){
		if (document.getElementById("img_" + pIndex).title=="Editar"){
			document.getElementById("img_" + pIndex).title = "Guardar";
			document.getElementById("img_" + pIndex).src = "../images/save-16.png";
			editControls(pIndex,true)
		}
		else{
			saveUpdate(document.getElementById("dtContable_" + pIndex).value,document.getElementById("nuCtaPte_" + pIndex).value,document.getElementById("ctg_" + pIndex).value, document.getElementById("nuCtaPteLbl_" + pIndex).innerHTML, document.getElementById("ctgLbl_" + pIndex).innerHTML, pIndex, "UPD")
			
		}
	}
	function saveUpdate(pDtContable, pCtaPte, pCTG, pCtaPteAnt, pCTGAnt, pIndex, pOption){
        ch.bind("informeCTGNoEnviados_Ajax.asp?pto=<% =g_strPuerto %>&option=" + pOption + "&dtContable=" + pDtContable + "&cartaPorte=" + pCtaPte + "&CTGAnt=" + pCTGAnt + "&cartaPorteAnt=" + pCtaPteAnt + "&CTG=" + pCTG, "saveUpdate_Callback(" + pIndex + ",'" + pOption + "')");
	    ch.send();
	}
	function saveUpdate_Callback(pIndex, pOption) {
	    var ret  = ch.response();
	    if (ret=="OK"){
			document.getElementById("img_" + pIndex).title = "Editar";
			document.getElementById("img_" + pIndex).src = "../images/edit-16x16.png";
			editControls(pIndex,false);
			document.getElementById("ctgLbl_" + pIndex).innerHTML = document.getElementById("ctg_" + pIndex).value;	    
			document.getElementById("nuCtaPteLbl_" + pIndex).innerHTML = document.getElementById("nuCtaPte_" + pIndex).value;
            document.getElementById("dtContableLbl_" + pIndex).innerHTML = document.getElementById("dtContable_" + pIndex).value;
			document.getElementById("div_" + pIndex).innerHTML = ""
			if (pOption=='DEL') submitInfo();
			}
	    else{
			document.getElementById("div_" + pIndex).innerHTML = "<img style='cursor:pointer;' src='../images/warning-16x16.png' title='" + ret + "'>"
	    }
	}
	function deleteItem(pIndex, pEnabled){
		saveUpdate("","",0, document.getElementById("nuCtaPteLbl_" + pIndex).innerHTML, document.getElementById("ctgLbl_" + pIndex).innerHTML, pIndex, "DEL")
	}
	
	function editControls(pIndex, pEnabled){
		var visibilityInput;
		var visibilityLabel;
		var positionLabel;
		if (pEnabled){
			visibilityInput = "visible";
			visibilityLabel = "hidden"
			positionLabel = "absolute"
		}	
		else{
			visibilityInput = "hidden";
			visibilityLabel = "visible"
			positionLabel = "relative"
		}	
		document.getElementById("ctg_" + pIndex).type = visibilityInput;
		document.getElementById("nuCtaPte_" + pIndex).type = visibilityInput;
		document.getElementById("dtContable_" + pIndex).type = visibilityInput;
		document.getElementById("ctgLbl_" + pIndex).style.position = positionLabel;
		document.getElementById("ctgLbl_" + pIndex).style.visibility = visibilityLabel;
		document.getElementById("nuCtaPteLbl_" + pIndex).style.position = positionLabel;
		document.getElementById("nuCtaPteLbl_" + pIndex).style.visibility = visibilityLabel;
		document.getElementById("dtContableLbl_" + pIndex).style.position = positionLabel;
		document.getElementById("dtContableLbl_" + pIndex).style.visibility = visibilityLabel;	
	}
</script>
</head>

<body>

<div id="toolbar"></div>

<form name="frmSel" id="frmSel">	
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> filtro - <%=GF_Traducir("CTGs No Informados")%> </h3>
    
    <div id="searchfilter" class="tableasidecontent">
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha") %> </div>    
        <div class="col16"> 
			<input type="text" name="fdesde" id="fdesde" size="15" readonly onClick="javascript:MostrarCalendario('fdesde', SeleccionarCalDesde)" value="<% =fdesde %>">
        </div>                        
        <span class="btnaction"><input type="submit" value="Buscar" id=submit1 name=submit1></span>
    </div>
</div><!-- END BUSCAR -->

<div class="col66"></div>
<% 
dim myEtapa
dim myEdit 
myEtapa = empty
	if hayError() then 
		Call showErrors() 
	else	%>
		<TABLE class="datagrid" id="TAB1" align="center" width="80%">
       <tbody>
			<%reg = 0
			while not rsGeneral.eof
				reg = reg + 1 
				if myEtapa <> rsGeneral("ETAPA") then
					myEtapa = rsGeneral("ETAPA")
					call printEncabezado(myEtapa)
				end if				
				%>
				<tr class="reg_Header_navdos">
					<TD align="center">
						<label id="dtContableLbl_<%=reg%>"><%=GF_FN2DTE(rsGeneral("DTCONTABLE"))%></label>
						<input maxlength="10" size="10" type="hidden" id="dtContable_<%=reg%>" value="<%=GF_FN2DTE(rsGeneral("DTCONTABLE"))%>">
					</TD>
					<TD align="center">
						<label id="nuCtaPteLbl_<%=reg%>"><%=left((trim(rsGeneral("NUCARTAPORTE"))),4) & "-" & mid((trim(rsGeneral("NUCARTAPORTE"))),5,8)%></label>
						<input maxlength="13" size="15" type="hidden" id="nuCtaPte_<%=reg%>" value="<%=left((trim(rsGeneral("NUCARTAPORTE"))),4) & "-" & mid((trim(rsGeneral("NUCARTAPORTE"))),5,8)%>">
					</TD>
					<TD align="center">
						<label id="ctgLbl_<%=reg%>"><%=rsGeneral("CTG")%></label>
						<input maxlength="8" size="9" type="hidden" id="ctg_<%=reg%>" value="<%=rsGeneral("CTG")%>">					
					</TD>
					<TD align="center"><%=rsGeneral("IDCAMION")%></TD>
					<TD align="center"><%=rsGeneral("CDPRODUCTO")%></TD>
					<TD align="center"><%=getEstadoDS(rsGeneral("CDESTADO"))%></TD>
					<TD align="center">					    
					    <%
					        if (rsGEneral("CDESTADO") <> CAMIONES_ESTADO_RECHAZADO) then
					            response.Write GF_EDIT_DECIMALS(rsGeneral("KILOSNETOS"), 0)
					        end if    
					    %>
					</TD>
					<td align="center"><% if ((myEtapa = 1) or not isNull(rsGeneral("ESTADOARRIBO"))) then response.Write  rsGeneral("ESTADOARRIBO") & "-" & getEstadoCTG(rsGeneral("ESTADOARRIBO"), rsGeneral("CDUSERARRIBO"))%> </td>
					<td align="center"><% if ((myEtapa = 1) or not isNull(rsGeneral("ESTADOCONFIRMACION"))) then response.Write rsGeneral("ESTADOCONFIRMACION") & "-" & getEstadoCTG(rsGeneral("ESTADOCONFIRMACION"), rsGeneral("CDUSERCONFIRMACION"))%> </td>
					<td align="center"><% if ((myEtapa = 1) or not isNull(rsGeneral("ESTADORECHAZO"))) then response.Write rsGeneral("ESTADORECHAZO") & "-" & getEstadoCTG(rsGeneral("ESTADORECHAZO"), rsGeneral("CDUSERRECHAZO"))%> </td>
					<td align="center"><% if ((myEtapa = 1) or not isNull(rsGeneral("ESTADODESVIO"))) then response.Write rsGeneral("ESTADODESVIO") & "-" & getEstadoCTG(rsGeneral("ESTADODESVIO"), rsGeneral("CDUSERDESVIO"))%> </td>						
					<%if rsGeneral("ETAPA") <> 5 then%>
						<td align="center">
							<span id="div_<%=reg%>"></span>
							<img id="img_<%=reg%>" style="cursor:pointer;" title="Editar" onclick="enabledEdit(<%=reg%>)" src="../images/edit-16x16.png">
							<img id="imgDel_<%=reg%>" style="cursor:pointer;" title="Quitar" onclick="deleteItem(<%=reg%>)" src="../images/delete-16x16.png">
						</td>
					<% else%>
						<!--<td align="center">&nbsp;</td>-->
					<%end if%>
				</tr>
			<%
				rsGeneral.movenext
			wend %>	  	
      </tbody>
           
	</TABLE>	
<%	end if%>
  	<%
	if (reg = 0) and not hayError() then		
		%>
		<tr>
			<td colspan="6"><% =GF_TRADUCIR("No se encontraron resultados.") %></td>
		</tr>
		<%
	end if 
	%>	
	<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">
	<input type="hidden" id="pto" name="pto" value="<% =g_strPuerto %>">	
	<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>">
	<div id="actionLabel" class="confirmsj" style="width:80%;visibility:hidden;"></div>
</form>
</body>
</html>
