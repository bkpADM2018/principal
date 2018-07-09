<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosSQL.asp"-->
<%
'----------------------------------------------------------------------------------------------------------------------
Function getEmbarques(pCdAviso, pCdProducto, pEstado, pFecha, pOrderBy)
	Dim strSQL,myWhere
	call buscarFiltrosEmbarques(myWhere,pCdAviso, pCdProducto, pEstado, pFecha)
	strSQL = " SELECT E.*,			"&_
			 "		  B.DSBUQUE     "&_			 
			 "  FROM EMBARQUES E	"&_
			 "	  INNER JOIN BUQUES B ON E.CDBUQUE = B.CDBUQUE " & myWhere & pOrderBy	
	call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)	
	Set getEmbarques = rs
End Function
'----------------------------------------------------------------------------------------------------------------------
Sub buscarFiltrosEmbarques(ByRef myWhere,pCdAviso, pCdProducto, pEstado, pFecha)
	if (pCdAviso > 0) then	Call mkWhere(myWhere, "E.CDAVISO", pCdAviso, "=", 1)
	'if (pCdProducto > 0) then Call mkWhere(myWhere, "ED.CDPRODUCTO", pCdProducto, "=", 1)
	if (pFecha <> "") then Call mkWhere(myWhere, "E.DTAVISO", pFecha, "=", 3)
End sub
'----------------------------------------------------------------------------------------------------------------------
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
'----------------------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto, myCdAvisoAnt, myOrderBy, myHtmlCargaAnterior, sqlCampoOrder, sqlTipoOrder, Conn, params
Dim lineasTotales, mostrar,paginaActual, fecAvisoA, fecAvisoM, fecAvisoD

g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)
g_cdAviso = GF_Parametros7("cdAviso",0,6)
call addParam("cdAviso", g_cdAviso, params)
g_cdProducto = GF_Parametros7("cmbCdProducto",0,6)
call addParam("cmbCdProducto", g_cdProducto, params)
g_estado = GF_Parametros7("estado",0,6)
call addParam("estado", g_estado, params)
sqlCampoOrder = GF_Parametros7("sqlCampoOrder", "", 6)
call addParam("sqlCampoOrder", sqlCampoOrder, params)
if sqlCampoOrder = "" then sqlCampoOrder = "E.CDAVISO"
sqlTipoOrder  = GF_Parametros7("sqlTipoOrder", "", 6)
call addParam("sqlTipoOrder", sqlTipoOrder, params)
if sqlTipoOrder = "" then sqlTipoOrder = "DESC"
fecAvisoD = GF_PARAMETROS7("fecAvisoD", "", 6)
call addParam("fecAvisoD", fecAvisoD, params)
fecAvisoM = GF_PARAMETROS7("fecAvisoM", "", 6)
Call addParam("fecAvisoM", fecAvisoM, params)
fecAvisoA = GF_PARAMETROS7("fecAvisoA", "", 6)
Call addParam("fecAvisoA", fecAvisoA, params)

if ((fecAvisoD <> "")and(fecAvisoM <> "")and(fecAvisoA <> "")) then 
	fechaAviso = fecAvisoA &"-"& fecAvisoM &"-"& fecAvisoD
end if
myOrderBy = " ORDER BY " & sqlCampoOrder & " " & sqlTipoOrder 
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

Set rsAvisos = getEmbarques(g_cdAviso, g_cdProducto, g_estado, fechaAviso, myOrderBy)

Call setupPaginacion(rsAvisos, paginaActual, mostrar)
lineasTotales = rsAvisos.recordcount


%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administracion de Embarques </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />	
	<link rel="stylesheet" href="../css/uploadmanager.css"	 type="text/css">
	<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>
</HEAD>
<script defer='defer' type='text/javascript' src="../Scripts/pngfix.js"></script>
<script type="text/javascript" src="../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/uploadManager.js"></script>
<script type="text/javascript" src="../scripts/JQueryUpload.js"></script>
<script type="text/javascript" src="../scripts/date.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/date.js"></script>
<script language="javascript">
	var ch = new channel();
	var up1;
	var myCdAviso;
	var valKilosBza;
	function onLoadPage(){
	<% 	if (not rsAvisos.eof) then %>
			var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "AdministracionMuelle.asp<% =params %>");						 
	<%	end if 	%>	    	
	}
		
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function verDetalle(pCdAviso){
		var pElement = document.getElementById("TBL_" + pCdAviso); 
		if (pElement.style.visibility == "hidden") {
		    loadDetalle(pCdAviso);
			pElement.style.visibility = "visible";
			pElement.style.position = "relative";
			document.getElementById("MAS_" + pCdAviso).src = "Images/menos.gif"
		}
		else{
			pElement.style.visibility = "hidden";
			pElement.style.position = "absolute";
			document.getElementById("MAS_" + pCdAviso).src = "Images/mas.gif"
			pElement.innerHTML = "<img src='images/Loading4.gif'>";
		}
	}
	
	function loadDetalle(pCdAviso){
		ch.bind("AdministracionMuelle_Ajax.asp?pto=<% =g_strPuerto %>&cdAviso=" + pCdAviso, "verDetalle_Callback(" + pCdAviso + ")");
	    ch.send();
	}
	function verDetalle_Callback(pCdAviso) {
	    var ret  = ch.response();
	    var pElement = document.getElementById("TBL_" + pCdAviso);
	    document.getElementById("TBL_" + pCdAviso).style.visibility = "visible";
	    pElement.innerHTML = ret;
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function saveDetails(pCdAviso){
		document.getElementById("SPAN_" + pCdAviso).style.visibility = "visible";
		document.getElementById("SPAN_" + pCdAviso).style.position = "relative";
		var kilosDS = 0;
		kilosDS = document.getElementById("kilosDraft_" +pCdAviso+ "_"+ document.getElementById("cdProducto_" + pCdAviso).value).value;		
		ch.bind("embarquesSaveDetails_AJAX.asp?Pto=<%=g_strPuerto%>&Aviso=" + pCdAviso + "&producto=" + document.getElementById("cdProducto_" + pCdAviso).value + "&kilos=" + document.getElementById("kilos_" + pCdAviso).value + "&cosecha=" + document.getElementById("cosecha_" + pCdAviso).value + "&permiso=" + document.getElementById("permiso_" + pCdAviso).value +"&kilosBza="+valKilosBza+"&kilosDraft="+kilosDS, "saveDetails_Callback(" + pCdAviso + ")");
		ch.send();			
	}
	function saveDetails_Callback(pCdAviso){
		document.getElementById("SPAN_" + pCdAviso).style.visibility = "hidden";
		document.getElementById("SPAN_" + pCdAviso).style.position = "absolute";
		var txt = ch.response();
		if (txt == "OK"){
			document.getElementById("TBL_" + pCdAviso).style.visibility = "hidden";
			loadDetalle(pCdAviso);
		} 
		else{
			document.getElementById("MSG_" + pCdAviso).innerHTML  = "<font color='red'>" + txt + "</font>";
		}
	}	
	/*function cargaEdit : carga los valores de los campos necesarios para cuando se quiera editar o agregar Cosecha 
						   A su vez resuelve el valor (Kilos) que tendra la Balanza a la hora de validarla (valKilosBza) */
	function cargaEdit(pCdAviso, pCdProducto, pDsProducto, pCdCosecha, pKilos, pPermiso){
		document.getElementById("MSG_" + pCdAviso).innerHTML = "";
		document.getElementById("cdProducto_" + pCdAviso).value = pCdProducto;				
		showHideFieldDraftSurvey(pCdAviso,false);
		document.getElementById("cosecha_" + pCdAviso).value = pCdCosecha;
		document.getElementById("spanProducto_" + pCdAviso).innerHTML = pCdProducto + " - " + pDsProducto;
		document.getElementById("kilos_" + pCdAviso).value = pKilos;		
		//Si se quiere asignar una cosecha y tiene un saldo de Kilos de Draft , tomo como valor defecto ese saldo Pendiente
		if((pCdCosecha == "")&&(document.getElementById("kilosDraft_"+pCdAviso+"_"+pCdProducto).value > 0)) document.getElementById("kilos_" + pCdAviso).value = document.getElementById("kilosDraftSinCosecha_"+pCdAviso+"_"+pCdProducto).value;
		document.getElementById("permiso_" + pCdAviso).value = pPermiso;		
		var kilosSinCosecha = document.getElementById("kilosSinCosecha_"+pCdAviso+"_"+pCdProducto).value;
		var kilosCosecha = document.getElementById("kilosCosecha_"+pCdAviso+"_"+pCdProducto).value;		
		if(pCdCosecha != ""){
			if(document.getElementById("kilosDraft_"+pCdAviso+"_"+pCdProducto).value > 0)
				valKilosBza = parseInt(kilosCosecha) - pKilos;				
			else
				valKilosBza = parseInt(kilosSinCosecha) + parseInt(pKilos);			
		}		
		else{				
			if(document.getElementById("kilosDraft_"+pCdAviso+"_"+pCdProducto).value > 0)
				valKilosBza = parseInt(kilosCosecha);				
			else
				valKilosBza = parseInt(pKilos);				
		}
		//Configuramos el Onclick del link Actualizar, para que valla a la configuracion inicial de cargar (saveDetails)
		$("#actulizar_"+ pCdAviso).each(function(){
			this.removeAttribute("onclick");
			this.setAttribute("onclick", "javascript:saveDetails("+pCdAviso+")");
		})		
	}
	function cancelDetails(pCdAviso){
		document.getElementById("spanProducto_" + pCdAviso).innerHTML = "";		
		document.getElementById("cosecha_" + pCdAviso).value = ""; 
		document.getElementById("kilos_" + pCdAviso).value = "";
		document.getElementById("permiso_" + pCdAviso).value = "";
		document.getElementById("MSG_" + pCdAviso).innerHTML = "";		
		showHideFieldDraftSurvey(pCdAviso, false)		
	}
	function cargaDel(pCdAviso, pCdProducto, pCdCosecha){
		if (confirm("Realmente desea quitar esta carga?")){
			ch.bind("embarquesSaveDetails_AJAX.asp?accion=DEL&Pto=<%=g_strPuerto%>&Aviso=" + pCdAviso + "&producto=" + pCdProducto + "&cosecha=" + pCdCosecha, "saveDetails_Callback(" + pCdAviso + ")");	
			ch.send();		
		}	
	}
		
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();
	}
	function setOrder(pCampo,pOrden){
		document.getElementById("sqlCampoOrder").value = pCampo;
		document.getElementById("sqlTipoOrder").value = pOrden;
		document.form1.submit();
	}	
	function asignarProducto(){
		document.getElementById("cdProducto").value = document.getElementById("cmbCdProducto").value;		
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
		
	function SeleccionarCalEmision(cal, date){		
		var str= new String(date);
		var fechaAviso = document.getElementById("fechaAviso_" + myCdAviso).value;
		var rtrn = compareDates(fechaAviso,"dd/MM/yyyy", str,"dd/MM/yyyy");		
		if (rtrn == 1){
			alert("La fecha del Draft Survey no puede ser menor a la fecha del Aviso!");
			$("issuedate_" + myCdAviso).val(<%= Left(session("Mmto"),8) %>);
			$("issuedateDiv_" + myCdAviso).val(<%= GF_FN2DTE(Left(session("Mmto"),8)) %>);			
		}
		else{
			str1 = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			document.getElementById("issuedate_" + myCdAviso).value = str1;
			document.getElementById("issuedateDiv_" + myCdAviso).innerHTML = str;				
		}		
		if (cal) cal.hide();
	}
	
	/* function newDraft : carga el Detalle para que se pueda ingresar los datos del Draft Survey, oculto los campos 
						   que no utilizo y muestro los nuevos (fecha y Archivo)	*/
	function newDraft(pCdAviso,pCdProducto,pDsProducto){		
		document.getElementById("spanProducto_" + pCdAviso).innerHTML = pCdProducto + " - " + pDsProducto;
		document.getElementById("cdProducto_" + pCdAviso).value = pCdProducto;				
		document.getElementById("kilos_" + pCdAviso).value = "";		
		showHideFieldDraftSurvey(pCdAviso, true);
		//asigno a la variable 'myCdAviso' el valor del cdAviso para luego utilizarla en la funcion 'SeleccionarCalEmision'
		myCdAviso = pCdAviso 
		up1 = new UploadHandler("dsFile_" + pCdAviso,"Temp");
		up1.draw();
		//modificamos el evento onclick de Actualizar, para que valla a otra funcion encargada de guardar el Draft		
		$("#actulizar_"+ pCdAviso).each(function(){			
			this.removeAttribute("onclick");			
			this.setAttribute("onclick", "javascript:saveDraftSurvey("+pCdAviso+","+pCdProducto+",0)");			
		})   		
	}
	
	/* function showHideFieldDraftSurvey : encargada de mostrar/ocultar los campos del draft Survey 
					      pFlagShowDraft : True(Muestra Draft, Oculta Cosecha) - False(Oculta Draft, Muestra Cosecha)	*/
	function showHideFieldDraftSurvey(pCdAviso, pFlagShowDraft){		
		if(pFlagShowDraft){
			document.getElementById("strCosecha_" + pCdAviso).innerHTML = "Fecha:";
			document.getElementById("cosecha_" + pCdAviso).type = "hidden"
			document.getElementById("issuedateLink_" + pCdAviso).style.visibility = 'visible';
			document.getElementById("issuedateDiv_" + pCdAviso).style.visibility = 'visible';
			document.getElementById("issuedateDiv_" + pCdAviso).style.position = 'relative';			
			document.getElementById("strPermiso_" + pCdAviso).innerHTML = "Archivo:";		
			document.getElementById("permiso_" + pCdAviso).type = 'hidden';
			document.getElementById("dsFile_" + pCdAviso).style.visibility = 'visible';		
			document.getElementById("dsFile_" + pCdAviso).style.position = 'relative';
			document.getElementById("trkgToepfer_" + pCdAviso).style.visibility = 'visible';		
			document.getElementById("trkgToepfer_" + pCdAviso).style.position = 'relative';
		}
		else{
			document.getElementById("strCosecha_" + pCdAviso).innerHTML = "Cosecha:";
			document.getElementById("cosecha_" + pCdAviso).type = "text"
			document.getElementById("issuedateLink_" + pCdAviso).style.visibility = 'hidden';
			document.getElementById("issuedateDiv_" + pCdAviso).style.visibility = 'hidden';
			document.getElementById("issuedateDiv_" + pCdAviso).style.position = 'absolute';			
			document.getElementById("strPermiso_" + pCdAviso).innerHTML = "Permiso:";
			document.getElementById("permiso_" + pCdAviso).type = 'text';
			document.getElementById("dsFile_" + pCdAviso).style.visibility = 'hidden';		
			document.getElementById("dsFile_" + pCdAviso).style.position = 'absolute';	
			document.getElementById("trkgToepfer_" + pCdAviso).style.visibility = 'hidden';		
			document.getElementById("trkgToepfer_" + pCdAviso).style.position = 'absolute';	
		}	
	}
	
	/* function saveDraftSurvey : Toma los datos del Detalle y los envia a la pagina cargaDraftSurvey_Ajax que los
								  valida y los graba	*/	
	function saveDraftSurvey(pCdAviso, pCdProducto, pIdDraft){			
		var cdProducto = document.getElementById("cdProducto_" + pCdAviso).value;
		var kilosBza = document.getElementById("kilosBza_" + pCdAviso + "_" + pCdProducto).value;		
		var kilosDraft = document.getElementById("kilos_" + pCdAviso).value;
		var kilosBzaToepfer = document.getElementById("kgToepfer_" + pCdAviso).value;
		var fechaDraft = document.getElementById("issuedate_" + pCdAviso).value;
		document.getElementById("dsFilePath_" + pCdAviso).value = up1.getFileName();		
		var archivo = up1.getFileName();
		var fechaBza = document.getElementById("fechaBza_" + pCdAviso + "_" + pCdProducto).value;
		var kilosCosecha = document.getElementById("kilosCosecha_"+pCdAviso+"_"+pCdProducto).value
		ch.bind("cargaDraftSurvey_Ajax.asp?Pto=<%=g_strPuerto%>&accion=<%=ACCION_GRABAR%>&cdAviso="+ pCdAviso +"&cdProducto=" + cdProducto + "&idDraft="+ pIdDraft +"&kilosDraft=" + kilosDraft + "&fechaDraft=" + fechaDraft + "&kilosBza=" + kilosBza + "&fechaBza="+ fechaBza +"&archivo="+ archivo +"&kilosCosecha="+kilosCosecha+"&kgBzaToepfer="+kilosBzaToepfer, "saveDraftSurvey_Callback(" + pCdAviso + ")");
		document.getElementById("SPAN_" + pCdAviso).style.visibility = "visible";
		document.getElementById("SPAN_" + pCdAviso).style.position = "relative";
		ch.send();
	}
	
	function saveDraftSurvey_Callback(pCdAviso){
		var ret  = ch.response();		
		if (ret == "OK"){
			loadDetalle(pCdAviso);
		} 
		else{
		    document.getElementById("SPAN_" + pCdAviso).style.visibility = "hidden";
		    document.getElementById("SPAN_" + pCdAviso).style.position = "absolute";
			document.getElementById("MSG_" + pCdAviso).innerHTML  = "<font color='red'>" + ret + "</font>";
		}
	}
	
	/*function cargaDelDraft : elimina un Draft Survey  */
	function cargaDelDraft(pCdAviso, pIdDraft){		
		if (confirm("Realmente desea quitar el Draft Survey?")){
			ch.bind("cargaDraftSurvey_Ajax.asp?accion=<%=ACCION_BORRAR%>&Pto=<%=g_strPuerto%>&idDraft="+pIdDraft, "saveDetails_Callback(" + pCdAviso + ")");
			ch.send();		
		}	
	}
	
	/*function cargaEditDraft : carga los valores de los campos necesarios para cuando se quiera editar un Draft Survey*/	
	function cargaEditDraft(pCdAviso, pIdDraft, pKilos, pCdProducto, pDsProducto,pfecha,pfechaFormat ){		
		document.getElementById("spanProducto_" + pCdAviso).innerHTML = pCdProducto + " - " + pDsProducto;
		document.getElementById("cdProducto_" + pCdAviso).value = pCdProducto;				
		document.getElementById("kilos_" + pCdAviso).value = pKilos;
		showHideFieldDraftSurvey(pCdAviso, true);
		document.getElementById("dsFile_" + pCdAviso).style.visibility = 'hidden';
		document.getElementById("strPermiso_" + pCdAviso).innerHTML = "";	
		document.getElementById("issuedateDiv_" + pCdAviso).innerHTML = pfechaFormat;
		document.getElementById("issuedate_" + pCdAviso).value = pfecha;	
		//asigno a la variable 'myCdAviso' el valor del cdAviso para luego utilizarla en la funcion 'SeleccionarCalEmision'
		myCdAviso = pCdAviso 
		up1 = new UploadHandler("dsFile_" + pCdAviso,"Temp");
		//Configuramos el Onclick del link Actualizar, para que valla a la configuracion inicial de cargar (saveDetails)
		$("#actulizar_"+ pCdAviso).each(function(){			
			this.removeAttribute("onclick");			
			this.setAttribute("onclick", "javascript:saveDraftSurvey("+pCdAviso+","+pCdProducto+","+pIdDraft+")");			
		})   
	}
	
</script>
<BODY onload="onLoadPage()">	
<form name="form1" id="form1" method=post>
<input type="hidden" name="sqlCampoOrder" id="sqlCampoOrder">
<input type="hidden" name="sqlTipoOrder" id="sqlTipoOrder">
<table border="0" cellpadding="0" cellspacing="0" width="95%" align="center">
  <tr><td><BR></BR></td></tr>
  <tr>
	  <td>
        <table id="tblBusqueda" width="95%" cellspacing="0" cellpadding="0" align="center" border="0">
			<tr>
			    <td width="8"><img src="../images/marco_r1_c1.gif"></td>
			    <td width="25%"><img src="../images/marco_r1_c2.gif" width="100%" height="8"></td>
			    <td width="8"><img src="../images/marco_r1_c3.gif"></td>
			    <td width="75%"><td>
			    <td></td>
			</tr>
			<tr>
			    <td width="8"><img src="../images/marco_r2_c1.gif"></td>
			    <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
			    <td width="8"><img src="../images/marco_r2_c3.gif"></td>
			    <td align="right">           		
			    </td>
			    <td></td>
			</tr>
			<tr>
			    <td><img src="../images/marco_r2_c1.gif" height="8"  width="8"></td>
			    <td></td>
			    <td><img src="../images/marco_c_s_d.gif" height="8" width="8"></td>
			    <td><img src="../images/marco_r1_c2.gif" width="100%" height="8"></td>
			    <td width="8"><img src="../images/marco_r1_c3.gif"></td>
			</tr>
			<tr>
			    <td height="100%"><img src="../images/marco_r2_c1.gif" height="100%" width="8"></td>
			    <td colspan="3">
                     <table width="100%" align="center" border="0">
                            <tr>
								<td width="12%" align="right"><% = GF_TRADUCIR("Aviso") %>:</td>
								<td width="38%">
									<input type="text" SIZE="3" MAXLENGTH="5" id="cdAviso" name="cdAviso" value="<% =g_cdAviso %>">
								</td>
								<td width="12%" align="right"><% = GF_TRADUCIR("Producto") %>:</td>
								<td width="38%">
									<select id="cmbCdProducto" name="cmbCdProducto" onchange="javascript:asignarProducto(this);">
										<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
										<%
										strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM PRODUCTOS ORDER BY DSPRODUCTO"
										call GF_BD_Puertos (g_strPuerto, rsProductos, "OPEN",strSQL)										
										while not rsProductos.eof 
											if cint(g_cdProducto) = cint(rsProductos("CDPRODUCTO")) then
												mySelected = "SELECTED"
											else
												mySelected = ""
											end if	%>
												<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
										<%	rsProductos.movenext
										wend
										%>							
									</select>
									<input type="hidden" id="cdProducto" name="cdProducto" value="<%=g_cdProducto%>">
								</td>                                
                            </tr>
                            <tr>
								<td width="12%" align="right"><% = GF_TRADUCIR("Fecha aviso") %>:</td>
								<td>
									<input type="text" size="1" maxLength="2" value="<% =fecAvisoD%>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAvisoD"> /
                                    <input type="text" size="1" maxLength="2" value="<% =fecAvisoM %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAvisoM"> /
                                    <input type="text" size="2" maxLength="4" value="<% =fecAvisoA %>" onKeyPress="return controlIngreso (this, event, 'N');" name="fecAvisoA">
                                </td>
                            </tr>
                            <tr>
								<td colspan="4" align="center">
									<input type="SUBMIT" value="Buscar..." id=cmdSearch name=cmdSearch onclick="submitInfo('<%=ACCION_SUBMITIR%>');">
								</td>	
                            </tr>								                            
                     </table>
	           </td>
	           <td height="100%"><img src="../images/marco_r2_c3.gif" width="8" height="100%"></td>
	         </tr>
	         <tr>
	           <td width="8"><img src="../images/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="../images/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="../images/marco_r3_c3.gif"></td>
	         </tr>
		 </table>
		</td>
	</tr>
	<tr><td><br></br></td></tr>
<% 	if (not rsAvisos.eof) then %>
	<tr>
       <td>    	
	  	   <table border=0 align="center" width="1024px" class="reg_Header" cellpadding=1 cellspacing=0>
	  			<tr>
					<td colspan=6>
						<div id="paginacion"></div>
					</td>
				</tr>		
				<tr class="reg_Header_nav">
					<td width="5%" align="center">.</td>
				    <td align="center">
						<img title="Ascendente" src="images/arrow_up_12x12.gif" onclick='setOrder("E.CDAVISO","ASC")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Aviso") %> &nbsp <img title="Descendente" src="images/arrow_down_12x12.gif" onclick='setOrder("E.CDAVISO","DESC")' style="cursor:pointer">		        
					</td>	
				    <td align="center">
						<img title="Ascendente" src="images/arrow_up_12x12.gif" onclick='setOrder("E.NUOPERACION","ASC")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Operacion") %> &nbsp <img title="Descendente" src="images/arrow_down_12x12.gif" onclick='setOrder("E.NUOPERACION","DESC")' style="cursor:pointer">		        
					</td>	
				    <td align="center">
						<img title="Ascendente" src="images/arrow_up_12x12.gif" onclick='setOrder("B.DSBUQUE","ASC")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Buque") %> &nbsp <img title="Descendente" src="images/arrow_down_12x12.gif" onclick='setOrder("B.DSBUQUE","DESC")' style="cursor:pointer">		        
						
					</td>	
				    <td align="center">
						<img title="Ascendente" src="images/arrow_up_12x12.gif" onclick='setOrder("E.DTAVISO","ASC")' style="cursor:pointer">&nbsp <% =GF_TRADUCIR("Fecha Aviso") %> &nbsp <img title="Descendente" src="images/arrow_down_12x12.gif" onclick='setOrder("E.DTAVISO","DESC")' style="cursor:pointer">		        
					</td>	
				</tr>
		<%		while not rsAvisos.EOF and (reg < mostrar) 		
					reg = reg + 1					
	    %>												
					<!--Imprimir linea de aviso nuevo -->
					<tr class="reg_Header_navdos" onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>
						<td align="center"><img id="MAS_<% =rsAvisos("CDAVISO") %>" onclick="verDetalle(<% =rsAvisos("CDAVISO") %>)" src='images/mas.gif'></td>
						<td align="center"><font size="2">
							<% =rsAvisos("CDAVISO") %></font>	
						</td>	
						<td align="center"><font size="2">
							<% =rsAvisos("NUOPERACION") %></font>	
						</td>	
						<td align="left"><font size="2">
							<% =rsAvisos("DSBUQUE") %> (<%=rsAvisos("CDBUQUE")%>)</font>	
						</td>	
						<td align="center">
							<font size="2">
								<% =GF_STANDARIZAR_FECHA_RTRN(rsAvisos("DTAVISO")) %>
							</font>	
							<input type="hidden" id="fechaAviso_<% =rsAvisos("CDAVISO") %>" name="fechaAviso_<% =rsAvisos("CDAVISO") %>" value="<% =GF_STANDARIZAR_FECHA_RTRN(rsAvisos("DTAVISO")) %>">
						</td>
					</tr>				
					<!--Imprimir Detalle Aviso -->					
					<tr>
						<td colspan="7"><div align="center" id="TBL_<% =rsAvisos("CDAVISO") %>" style="visibility: hidden; position:absolute; "><img src='images/Loading4.gif'></div></td>
					</tr>								
					<!--Fin Imprimir Detalle Aviso -->
			    	<!--Fin imprimir linea de aviso -->
<%                  rsAvisos.movenext
		        wend
	end if
%>
<input type="hidden" name="accion" id="accion" value="<%= accion %>">
</form>
</BODY>
</HTML>