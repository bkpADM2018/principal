<!--#include file="cartadeporteEditCommon.asp"-->
<%

    Call initTaskAccessInfo(TASK_POS_MODIFICACION_HISTORICA, session("DIVISION_PUERTO"))

    if (not isFormSubmit) then Set g_rs = loadDataCartaPorte(g_ctaPte, g_idCamion, g_dtContable, g_pto)
    Call fetchCabeceraCartaPorte()
    Call fetchIntervinientesCartaPorte()
    Call fecthProductoCartaPorte()
    Call fecthTransporteCartaPorte()
    Call fecthDescargaCartaPorte()
    if (accion = ACCION_CONTROLAR)or(accion = ACCION_GRABAR) then
        errorCabecera      = controlarCabeceraCartaPorte()
        errorInterviniente = controlarIntervinientesCartaPorte()
        errorProducto      = controlarProductoCartaPorte()
        errorTransporte    = controlarTransporteCartaPorte()
        errorDescarga      = controlarDescargaCartaPorte()
        if ((not hayErroresCartaPorte(g_Error))and(accion = ACCION_GRABAR)) then
            Call inicializarLogCartaPorte()
            Call grabarCabeceraCtaPte()
            Call grabarIntervinienteCtaPte()
            Call grabarProductoCtaPte()
            Call grabarTransporteCtaPte()
            Call grabarDescargaCtaPte()
            if (oDiccModificaciones.Count <> 0) then
                Call agregarRegistrosAuditoria(g_IdCamion,g_dtContable,g_strPuerto)
                auxNroAjuste = actualizarAjusteCamion(g_IdCamion,g_dtContable,auxGrano,auxGranoOld,auxDestinatario,auxDestinatarioOld,auxMerma,auxMermaOld,auxPesadaBruto,auxPesadaBrutoOld,auxPesadaTara,auxPesadaTaraOld)
                Call rearmarCartaPorte(g_IdCamion,g_dtContable,auxGrano,auxGranoOld,auxDestinatario,auxDestinatarioOld,auxMerma,auxMermaOld,auxPesadaBruto,auxPesadaBrutoOld,auxPesadaTara,auxPesadaTaraOld)
                Call rearmarStockFisico(g_IdCamion,g_dtContable,auxGrano,auxGranoOld,auxDestinatario,auxDestinatarioOld,auxMerma,auxMermaOld,auxPesadaBruto,auxPesadaBrutoOld,auxPesadaTara,auxPesadaTaraOld)
                'Una vez modificado los Ajustes del Camion igualo las variables afectadas a este cambio.
                auxDestinatarioOld = auxDestinatario
                auxDestinatarioCuitOld = Trim(auxDestinatarioCuit1) & Trim(auxDestinatarioCuit2) & Trim(auxDestinatarioCuit3)
                auxGranoOld = auxGrano
                auxPesadaBrutoOld = auxPesadaBruto
                auxPesadaTaraOld = auxPesadaTara
                Call enviarMailCtaPte()
                flagGrabo = true
            end if
        end if
    end if
'-----------------------------------------------------------------------------------------------------------------------------------
'TAREA 1780
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Poseidon - Modificaci&oacuten hist&oacuterica de camiones </title>


<link rel="stylesheet" type="text/css" href="../css/ActiSAIntra-1.css">	
<link rel="stylesheet" type="text/css" href="../css/main.css">
<link rel="stylesheet" type="text/css" href="../css/toolbar.css">
<link rel="stylesheet" type="text/css" href="../css/Header.css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<style>
	input[type=text] { text-align:right;}
	input[type="radio"] { float:right;}
  	label { padding: 0; marign: 0; display: block; }
  	textarea { width: 100%; border: 0px; padding: 0px; }
  	.the-fix { }
	.celda {
		border-radius:8px 8px 8px 8px;
	}
	#barratab	{
		/*padding: 6px;*/
		padding-bottom: 4px;
		padding-top: 4px;
		font-size: 16px;
		color: #FFF;
		font-weight:bold;
		width: 100%;
		text-align:center;
		vertical-align: middle;
		background: #006b4a;
		}
	.codbar	{
		/*border: 1px solid #FFF;*/
		padding: 1px;
		font-size: 10px;
		color: #FFF;
		font-weight:bold;
		width: 100%;
		text-align:center;
		vertical-align:bottom;
		background: #006b4a;
	}
	/*#container {
    	position:relative;
    	width:95%;
    	height:auto;
	}*/
	.firmas {
		border: 1px solid #999;
		border-radius: 8px;
		color: #333;
		/*margin-left: 6px;
		margin-right: 6px;*/
		padding-top: 50px;
		padding-bottom: 8px;
		/*padding-left: 28px;
		padding-right: 28px;
		background: rgba(256, 256, 256, 1);*/
	}
	#w1 {
    	display: inline-table;
	}
	.firmas > p {
    	display: table-cell;
	}
 </style>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/date.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript">
    var ch = new channel();
    var valProvincia;
	function bodyOnLoad(){
	    var tb = new Toolbar('toolbar');
	    tb.addButton("toolbar-save", "Guardar", "saveCtaPte()");
	    tb.addButton("toolbar-control", "Controlar", "controlCtaPte()");
	    tb.addButton("../../images/DatosCalidad-16x16.png", "Datos de Calidad", "updateCalidad()");
	    tb.draw();
        
        //DatosCalidad-16x16.png
        <% if (flagGrabo) then %>
            var strMsj = "Se ha guardado correctamente.";
	        document.getElementById("respuestaCtaPte").className = "reg_Header_success";
	        document.getElementById("respuestaCtaPte").style.width = "100%";
	        <% if (Cdbl(auxNroAjuste) <> 0) then %>
                strMsj = strMsj + "El numero de ajuste de modificacion es <%=auxNroAjuste%>";
            <% end if%>    
            document.getElementById("respuestaCtaPte").innerHTML = strMsj;
        <% else %>
            if ('<%=g_Error %>' != "") {
	            document.getElementById("respuestaCtaPte").innerHTML = "<%=g_Error %>";
	            document.getElementById("respuestaCtaPte").className = "reg_Header_Error";
	            document.getElementById("respuestaCtaPte").style.width = "100%";
            }
        <% end if %>
	    loadAutocomplete();
	}
	function loadAutocomplete() {
	    autoCompleteProcedencia();
    }
    function updateCalidad() {
        myPopUp = new winPopUp('Iframe', 'cartadeportePopUpRubros.asp?pto=<%=g_pto %>&dtContable=<%=g_dtContable%>&ctaPte=<%=g_ctaPte %>&idCamion=<%=g_IdCamion %>&cdProducto=<%=auxGranoOld %>&mermaOld=<%= auxMermaOld%>&destinatario=<%=auxDestinatario %>', '700', '700', 'Informacion de Camiones', 'resetForm()');
	}
	function dimensionarIframe(p_width, p_height) {
	    myPopUp.resize(p_width, p_height);
	}
    function resetForm() {
		document.getElementById("accion").value = ''
        //submitInfo();
	}
	function submitInfo() {
		document.getElementById("frmSel").submit();
	}
	function saveCtaPte() {
		document.getElementById("accion").value = '<%=ACCION_GRABAR%>'
        submitInfo();
	}
	function controlCtaPte() {
	    document.getElementById("accion").value = '<%=ACCION_CONTROLAR%>'
	    submitInfo();
	}
	function MostrarCalendario(p_objID, funcSel) {
	    var dte = new Date();
	    var elem = document.getElementById(p_objID);
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
	    var str = new String(date);
	    var myObj;
	    document.getElementById("issuedateCargaTXT").value = str;
	    document.getElementById("issuedateCarga").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
	    if (cal) cal.hide();
	}
	function SeleccionarCalVencimiento(cal, date) {
	    var str = new String(date);
	    var myObj;
	    document.getElementById("issuedateVencimientoTXT").value = str;
	    document.getElementById("issuedateVencimiento").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
	    if (cal) cal.hide();
	}
	function CerrarCal(cal) {
	    cal.hide();
	}
	
	function controlarCUIT(pTipoInterviniente) {	    
	    var cuit1 = document.getElementById("Cuit_" + pTipoInterviniente + "_1").value;
		var	cuit2 = document.getElementById("Cuit_" + pTipoInterviniente + "_2").value;
		var	cuit3 = document.getElementById("Cuit_" + pTipoInterviniente + "_3").value;
		var cuit1Old = document.getElementById("CuitOld_" + pTipoInterviniente + "_1").value;
		var cuit2Old = document.getElementById("CuitOld_" + pTipoInterviniente + "_2").value;
		var cuit3Old = document.getElementById("CuitOld_" + pTipoInterviniente + "_3").value;		
        if ((cuit1 != cuit1Old) || (cuit2 != cuit2Old) || (cuit3 != cuit3Old)) {
			var myCuit = cuit1 + cuit2 + cuit3;
			var myDs = document.getElementById("valDs_" + pTipoInterviniente).value;
			var myCd = document.getElementById("valCd_" + pTipoInterviniente).value;					
			ch.bind("cartadeporteEditAjax.asp?pto=<%=g_pto %>&cuit=" + myCuit + "&tipoInterviniente=" + pTipoInterviniente + "&ds=" + myDs + "&cd=" + myCd +"&accion=<%=ACCION_CONTROLAR%>", "controlarCUIT_callBack(" + pTipoInterviniente + ")");
			ch.send();
	    }		
    }
    function limpiarCuit(pTipoInterviniente) {
        document.getElementById("valDs_" + pTipoInterviniente).value = "No se encontraron resultados";
        document.getElementById("div_" + pTipoInterviniente).innerHTML = "No se encontraron resultados";
        document.getElementById("valCd_" + pTipoInterviniente).value = 0;       
		document.getElementById("CuitOld_" + pTipoInterviniente + "_1").value = 0;
		document.getElementById("CuitOld_" + pTipoInterviniente + "_2").value = 0;
		document.getElementById("CuitOld_" + pTipoInterviniente + "_3").value = 0;
    }
	function controlarCUIT_callBack(pTipoInterviniente) {
	    var resp = ch.response();
	    if (resp.indexOf("|") != -1) {
            //Un interviniente
	        var vals = resp.split("|");
	        document.getElementById("valDs_" + pTipoInterviniente).value = vals[1];	 
	        document.getElementById("valCd_" + pTipoInterviniente).value = vals[0];	        
			resp = vals[1];
	    }
	    else {
	        if (resp == "<%=ESTADO_BAJA %>") {
	            //No encontro nada
				limpiarCuit(pTipoInterviniente);
	            resp = document.getElementById("valDs_" + pTipoInterviniente).value;
            }
	    }	    
	    document.getElementById("div_" + pTipoInterviniente).innerHTML = resp;
		document.getElementById("CuitOld_" + pTipoInterviniente + "_1").value = document.getElementById("Cuit_" + pTipoInterviniente + "_1").value;
		document.getElementById("CuitOld_" + pTipoInterviniente + "_2").value = document.getElementById("Cuit_" + pTipoInterviniente + "_2").value;
		document.getElementById("CuitOld_" + pTipoInterviniente + "_3").value = document.getElementById("Cuit_" + pTipoInterviniente + "_3").value;

	}
    function changeDsCUIT(e, p_TipoInterviniente) {
        var strValue = e.value;
        var vals = strValue.split("-");
        document.getElementById("valDs_" + p_TipoInterviniente).value = vals[1];  
            document.getElementById("valCd_" + p_TipoInterviniente).value = vals[0];
        
    }    
    function changeProducto(e) {
        var myProd = e.value;
        ch.bind("cartadeporteEditAjax.asp?pto=<%=g_pto %>&cdProducto=" + myProd + "&accion=<%=ACCION_VISUALIZAR%>", "changeProducto_callBack()");
        ch.send();
    }
    function changeProducto_callBack() {
        var resp = ch.response();
        var select = document.getElementById("cmbBiotecnologia");
        var length = select.options.length;
        for (var i = 1; i < length; i++) {
            select.remove(i);
        }
        if (resp != "") $('#cmbBiotecnologia').append(resp);

    }
    
    function autoCompleteProcedencia() {
        $("#procedenciaDs").autocomplete({
            minLength: 3,
            source: "puertosStreamElementos.asp?tipo=JQProcedenciasPto&pto=<%=g_pto%>&pcia=" + document.getElementById("cmbProvincia").value,
            focus: function (event, ui) {
                $("#procedenciaDs").val(ui.item.dsprocedencia);
                return false;
            },
            select: function (event, ui) {
                $("#procedenciaDs").val(ui.item.dsprocedencia);
                $("#procedenciaCd").val(ui.item.cdprocedencia);
                return false;
            },
            change: function (event, ui) {
                if (!ui.item) {
                    $("#procedenciaDs").val("");
                    $("#procedenciaCd").val(0);
                }
            }
        })
		.data("autocomplete")._renderItem = function (ul, item) {
		    return $("<li></li>")
				.data("item.autocomplete", item)
				.append("<a>" + item.cdprocedencia + " - <font style='font-size:10;'>" + item.dsprocedencia + "</font></a>")
				.appendTo(ul);
		};
    }
    function changeProvincia(e) {
        loadAutocomplete();
        document.getElementById("procedenciaDs").value = "";
        document.getElementById("procedenciaCd").value = 0;
    }
    function calcularPeso(e){
        var bruto = document.getElementById("pesoBruto").value;
        var tara  = document.getElementById("pesoTara").value;
        var neto  = document.getElementById("pesoNeto").value;
        if (e.name != "pesoNeto") {
            document.getElementById("pesoNeto").value = parseInt(bruto) - parseInt(tara);
        }
        else {
            if (bruto != 0) {
                document.getElementById("pesoTara").value = parseInt(bruto) - parseInt(e.value);
            }
            else {
                if (tara != 0) document.getElementById("pesoBruto").value = parseInt(e.value) + parseInt(tara);
            }
        }
    }
    function controlarDiguitosCosecha(evento) {
        var ascii = (document.all) ? evento.keyCode : evento.which;
        if ((ascii >= 48 && ascii <= 57) || (ascii == 127) || (ascii == 8) || (ascii == 0)) {
            return true;
        }
        else {
            return false;
        }
    }
    function calcularPesada(){
        var bruto = document.getElementById("pesadaBruto").value;
        var tara  = document.getElementById("pesadaTara").value;
        var total = parseInt(bruto) - parseInt(tara);
        document.getElementById("pesadaNeto").value = total;
        document.getElementById("divPesadaNeto").innerHTML = total;
    }
</script>
</head>
<body onload="bodyOnLoad()">
<div id="toolbar"></div>
<form id="frmSel" name="frmSel" method="post">
    <div class="col66"></div>
    <div class="tableaside size100">
    	<h3> <% =GF_TRADUCIR("EDITAR CARTA DE PORTE") %> </h3>
    </div>
    <table width="95%" align="center"><tr><td><div id="respuestaCtaPte" ></div></td></tr></table>
    <table width="95%" border="1" bordercolor="#006b4a" align="center" cellpadding="0" cellspacing="0">
	    <tr>
    	    <td rowspan="3"><img src="../images/Header_CartaPorte.png"/></td>
            <td width="50%" height="30" align="center" valign="bottom"><div class="codbar">Carta de Porte</div>
            <p><input type="" id="txtCartaPorte1" name="txtCartaPorte1" value="<%= auxCartaPorte1%>" maxlength="4" size="2" />-
               <input type="" id="txtCartaPorte2" name="txtCartaPorte2" value="<%= auxCartaPorte2%>" maxlength="8" size="8" />
            </td>
            <td width="50%" height="50" align="center" valign="bottom"><div class="codbar">C.E.E.</div>
            <p><input type="" value="" size="5" readonly /> - <input type="" value="" size="15" readonly />
            </td>
	    </tr>
	    <tr>
	      <td rowspan="2" align="center" style="font-size:14px; font-weight:bold">C.T.G:
            <input type="" id="txtCTG" name="txtCTG" value="<%=auxCTG %>" maxlength="10" size="20" onKeyPress="return controlIngreso (this, event, 'N');" />
            <input type="hidden" id="txtCTGOld" name="txtCTGOld" value="<%=auxCTGOld %>" />
          </td>
	      <td align="right">FECHA DE CARGA: &nbsp;
            <a id="dtFechaCarga" href="javascript:MostrarCalendario('imgDtFechaCarga', SeleccionarCalEmision)">
			    <img id="imgDtFechaCarga" src="../images/calendar-16.png">
			</a>
            <input id="issuedateCargaTXT" name="issuedateCargaTXT" type="" value="<%=GF_FN2DTE(auxDtCartaPorte)%>" size="10" readonly="readonly" />
            <input type="hidden" id="issuedateCarga" name="issuedateCarga" value="<%=auxDtCartaPorte%>"/>
            <input type="hidden" id="issuedateCargaOld" name="issuedateCargaOld" value="<%=auxDtCartaPorteOld%>"/>
          </td>
      </tr>
	    <tr>
	      <td align="right">FECHA DE VENCIMIENTO: &nbsp;
	        <a id="dtFechaVencimiento" href="javascript:MostrarCalendario('imgDtFechaVencimiento', SeleccionarCalVencimiento)">
			    <img id="imgDtFechaVencimiento" src="../images/calendar-16.png">
			</a>
            <input id="issuedateVencimientoTXT" name="issuedateVencimientoTXT" type="" value="<%=GF_FN2DTE(auxDtVencimiento)%>" size="10" readonly=readonly />
            <input type="hidden" id="issuedateVencimiento" name="issuedateVencimiento" value="<%=auxDtVencimiento%>"/>
            <input type="hidden" id="issuedateVencimientoOld" name="issuedateVencimientoOld" value="<%=auxDtVencimientoOld%>"/>
          </td>
      </tr>
	    <tr>
	      <td colspan="3"><div id="barratab">CARTA DE PORTE PARA EL TRANSPORTE AUTOMOTOR DE GRANOS</div></td>
      </tr>
    </table>
    <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="10px" align="center">1</th>
                <td align="left" colspan="4" style="border-top:medium; border-top-right-radius: 8px;" height="18px">&nbsp; DATOS DE INTERVINIENTES EN EL TRASLADO DE GRANO</td>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td colspan="2" align="left" style="color:#000">TITULAR CARTA PORTE</td>
                <td align="left" width="50%">
                    <div id="div_<%=INTERVINIENTE_TITULAR%>"><%= auxTitularDs %></div>
                    <input type="hidden" id="valDs_<%=INTERVINIENTE_TITULAR%>" name="valDs_<%=INTERVINIENTE_TITULAR%>" value="<%=auxTitularDs %>"/>
                    <input type="hidden" id="valCd_<%=INTERVINIENTE_TITULAR%>" name="valCd_<%=INTERVINIENTE_TITULAR%>" value="<%=auxTitularCd %>"/>
                    <input type="hidden" id="valCdOld_<%=INTERVINIENTE_TITULAR%>" name="valCdOld_<%=INTERVINIENTE_TITULAR%>" value="<%=auxTitularCdOld %>"/>
                    <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_TITULAR%>" name="valCuitOld_<%=INTERVINIENTE_TITULAR%>" value="<%=auxTitularCuitOld %>"/>
                </td>
                <td width="10%" align="center" style="color:#000">CUIT Nro</td>
                <td width="20%" align="left">
                    <input type="text" size="1" maxlength="2" id="Cuit_<%=INTERVINIENTE_TITULAR%>_1" name="Cuit_<%=INTERVINIENTE_TITULAR%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TITULAR%>);" oncontextmenu="return false" value="<%=auxTitularCuit1 %>" />-
                    <input type="text" size="8" maxlength="8" id="Cuit_<%=INTERVINIENTE_TITULAR%>_2" name="Cuit_<%=INTERVINIENTE_TITULAR%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TITULAR%>);" oncontextmenu="return false" value="<%=auxTitularCuit2 %>" />-
                    <input type="text" size="1" maxlength="1" id="Cuit_<%=INTERVINIENTE_TITULAR%>_3" name="Cuit_<%=INTERVINIENTE_TITULAR%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TITULAR%>);" oncontextmenu="return false" value="<%=auxTitularCuit3 %>" />					
					<input type="hidden" id="CuitOld_<%= INTERVINIENTE_TITULAR%>_1" name="CuitOld_<%= INTERVINIENTE_TITULAR%>_1" value="<%=auxTitularCuit1 %>"  />
					<input type="hidden" id="CuitOld_<%= INTERVINIENTE_TITULAR%>_2" name="CuitOld_<%= INTERVINIENTE_TITULAR%>_2" value="<%=auxTitularCuit2 %>"  />
					<input type="hidden" id="CuitOld_<%= INTERVINIENTE_TITULAR%>_3" name="CuitOld_<%= INTERVINIENTE_TITULAR%>_3" value="<%=auxTitularCuit3 %>"  />
                </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">INTERMEDIARIO</td>
              <td align="left">
                 <div id="div_<%= INTERVINIENTE_INTERMEDIARIO%>"><%=auxIntermediarioDs %></div>
                 <input type="hidden" id="valDs_<%= INTERVINIENTE_INTERMEDIARIO%>" name="valDs_<%= INTERVINIENTE_INTERMEDIARIO%>" value="<%= auxIntermediarioDs%>" />
                 <input type="hidden" id="valCd_<%= INTERVINIENTE_INTERMEDIARIO%>" name="valCd_<%= INTERVINIENTE_INTERMEDIARIO%>" value="<%= auxIntermediarioCd%>" />
                 <input type="hidden" id="valCdOld_<%=INTERVINIENTE_INTERMEDIARIO%>" name="valCdOld_<%=INTERVINIENTE_INTERMEDIARIO%>" value="<%=auxIntermediarioCdOld %>"/>
                 <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_INTERMEDIARIO%>" name="valCuitOld_<%=INTERVINIENTE_INTERMEDIARIO%>" value="<%=auxIntermediarioCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_1" name="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_INTERMEDIARIO%>);" oncontextmenu="return false" value="<%=auxIntermediarioCuit1 %>"/>-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_2" name="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_INTERMEDIARIO%>);" oncontextmenu="return false" value="<%=auxIntermediarioCuit2 %>"/>-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_3" name="Cuit_<%= INTERVINIENTE_INTERMEDIARIO%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_INTERMEDIARIO%>);" oncontextmenu="return false" value="<%=auxIntermediarioCuit3 %>"/>
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_1" name="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_1" value="<%=auxIntermediarioCuit1 %>"  />
                  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_2" name="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_2" value="<%=auxIntermediarioCuit2 %>"  />
                  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_3" name="CuitOld_<%= INTERVINIENTE_INTERMEDIARIO%>_3" value="<%=auxIntermediarioCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">REMITENTE COMERCIAL</td>
              <td align="left">
                 <div id="div_<%= INTERVINIENTE_REMITENTE%>"><%=auxRemitenteDs %></div>
                 <input type="hidden" id="valDs_<%= INTERVINIENTE_REMITENTE%>" name="valDs_<%= INTERVINIENTE_REMITENTE%>" value="<%= auxRemitenteDs%>"/>
                 <input type="hidden" id="valCd_<%= INTERVINIENTE_REMITENTE%>" name="valCd_<%= INTERVINIENTE_REMITENTE%>" value="<%= auxRemitenteCd%>"/>
                 <input type="hidden" id="valCdOld_<%=INTERVINIENTE_REMITENTE%>" name="valCdOld_<%=INTERVINIENTE_REMITENTE%>" value="<%=auxRemitenteCdOld%>"/>
                 <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_REMITENTE%>" name="valCuitOld_<%=INTERVINIENTE_REMITENTE%>" value="<%=auxRemitenteCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_REMITENTE%>_1" name="Cuit_<%= INTERVINIENTE_REMITENTE%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_REMITENTE%>);" oncontextmenu="return false" value="<%=auxRemitenteCuit1 %>"  />-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_REMITENTE%>_2" name="Cuit_<%= INTERVINIENTE_REMITENTE%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_REMITENTE%>);" oncontextmenu="return false" value="<%=auxRemitenteCuit2 %>"  />-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_REMITENTE%>_3" name="Cuit_<%= INTERVINIENTE_REMITENTE%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_REMITENTE%>);" oncontextmenu="return false" value="<%=auxRemitenteCuit3 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_REMITENTE%>_1" name="CuitOld_<%= INTERVINIENTE_REMITENTE%>_1" value="<%=auxRemitenteCuit1 %>"  />
                  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_REMITENTE%>_2" name="CuitOld_<%= INTERVINIENTE_REMITENTE%>_2" value="<%=auxRemitenteCuit2 %>"  />
                  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_REMITENTE%>_3" name="CuitOld_<%= INTERVINIENTE_REMITENTE%>_3" value="<%=auxRemitenteCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">CORREDOR</td>
              <td align="left">
                  <div id="div_<%= INTERVINIENTE_CORREDOR%>"><%=auxCorredorDs %></div>
                  <input type="hidden" id="valDs_<%= INTERVINIENTE_CORREDOR%>" name="valDs_<%= INTERVINIENTE_CORREDOR%>" value="<%= auxCorredorDs%>"/>
                  <input type="hidden" id="valCd_<%= INTERVINIENTE_CORREDOR%>" name="valCd_<%= INTERVINIENTE_CORREDOR%>" value="<%= auxCorredor%>"/>
                  <input type="hidden" id="valCdOld_<%=INTERVINIENTE_CORREDOR%>" name="valCdOld_<%=INTERVINIENTE_CORREDOR%>" value="<%=auxCorredorOld%>"/>
                  <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_CORREDOR%>" name="valCuitOld_<%=INTERVINIENTE_CORREDOR%>" value="<%=auxCorredorCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_CORREDOR%>_1" name="Cuit_<%= INTERVINIENTE_CORREDOR%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CORREDOR%>);" oncontextmenu="return false" value="<%=auxCorredorCuit1 %>" />-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_CORREDOR%>_2" name="Cuit_<%= INTERVINIENTE_CORREDOR%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CORREDOR%>);" oncontextmenu="return false" value="<%=auxCorredorCuit2 %>" />-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_CORREDOR%>_3" name="Cuit_<%= INTERVINIENTE_CORREDOR%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CORREDOR%>);" oncontextmenu="return false" value="<%=auxCorredorCuit3 %>" />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CORREDOR%>_1" name="CuitOld_<%= INTERVINIENTE_CORREDOR%>_1" value="<%=auxCorredorCuit1 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CORREDOR%>_2" name="CuitOld_<%= INTERVINIENTE_CORREDOR%>_2" value="<%=auxCorredorCuit2 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CORREDOR%>_3" name="CuitOld_<%= INTERVINIENTE_CORREDOR%>_3" value="<%=auxCorredorCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">REPRESENTANTE/ENTREGADOR</td>
              <td align="left">
                  <div id="div_<%= INTERVINIENTE_ENTREGADOR%>"><%=auxEntregadorDs %></div>
                  <input type="hidden" id="valDs_<%= INTERVINIENTE_ENTREGADOR%>" name="valDs_<%= INTERVINIENTE_ENTREGADOR%>" value="<%= auxEntregadorDs%>"/>
                  <input type="hidden" id="valCd_<%= INTERVINIENTE_ENTREGADOR%>" name="valCd_<%= INTERVINIENTE_ENTREGADOR%>" value="<%= auxEntregador%>"/>
                  <input type="hidden" id="valCdOld_<%=INTERVINIENTE_ENTREGADOR%>" name="valCdOld_<%=INTERVINIENTE_ENTREGADOR%>" value="<%=auxEntregadorOld%>"/>
                  <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_ENTREGADOR%>" name="valCuitOld_<%=INTERVINIENTE_ENTREGADOR%>" value="<%=auxEntregadorCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_1" name="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_ENTREGADOR%>);" oncontextmenu="return false" value="<%=auxEntregadorCuit1 %>"/>-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_2" name="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_ENTREGADOR%>);" oncontextmenu="return false" value="<%=auxEntregadorCuit2 %>"/>-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_3" name="Cuit_<%= INTERVINIENTE_ENTREGADOR%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_ENTREGADOR%>);" oncontextmenu="return false" value="<%=auxEntregadorCuit3 %>"/>
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_1" name="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_1" value="<%=auxEntregadorCuit1 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_2" name="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_2" value="<%=auxEntregadorCuit2 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_3" name="CuitOld_<%= INTERVINIENTE_ENTREGADOR%>_3" value="<%=auxEntregadorCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">DESTINATARIO</td>
              <td align="left">
                   <div id="div_<%= INTERVINIENTE_DESTINATARIO%>"><%=auxDestinatarioDs %></div>
                   <input type="hidden" id="valDs_<%= INTERVINIENTE_DESTINATARIO%>" name="valDs_<%= INTERVINIENTE_DESTINATARIO%>" value="<%= auxDestinatarioDs%>"/>
                   <input type="hidden" id="valCd_<%= INTERVINIENTE_DESTINATARIO%>" name="valCd_<%= INTERVINIENTE_DESTINATARIO%>" value="<%= auxDestinatario%>"/>
                   <input type="hidden" id="valCdOld_<%=INTERVINIENTE_DESTINATARIO%>" name="valCdOld_<%=INTERVINIENTE_DESTINATARIO%>" value="<%=auxDestinatarioOld%>"/>
                   <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_DESTINATARIO%>" name="valCuitOld_<%=INTERVINIENTE_DESTINATARIO%>" value="<%=auxDestinatarioCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                   <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_1" name="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_DESTINATARIO%>);" oncontextmenu="return false" value="<%=auxDestinatarioCuit1 %>" />-
                   <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_2" name="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_DESTINATARIO%>);" oncontextmenu="return false" value="<%=auxDestinatarioCuit2 %>" />-
                   <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_3" name="Cuit_<%= INTERVINIENTE_DESTINATARIO%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_DESTINATARIO%>);" oncontextmenu="return false" value="<%=auxDestinatarioCuit3 %>" />
				   <input type="hidden" id="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_1" name="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_1" value="<%=auxDestinatarioCuit1 %>"  />
				   <input type="hidden" id="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_2" name="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_2" value="<%=auxDestinatarioCuit2 %>"  />
				   <input type="hidden" id="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_3" name="CuitOld_<%= INTERVINIENTE_DESTINATARIO%>_3" value="<%=auxDestinatarioCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">DESTINO</td>
              <td align="left">&nbsp;</td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">TRANSPORTISTA</td>
              <td align="left">
                  <div id="div_<%= INTERVINIENTE_TRANSPORTISTA%>"><%=auxTransportistaDs %></div>
                  <input type="hidden" id="valDs_<%= INTERVINIENTE_TRANSPORTISTA%>" name="valDs_<%= INTERVINIENTE_TRANSPORTISTA%>" value="<%=auxTransportistaDs %>"/>
                  <input type="hidden" id="valCd_<%= INTERVINIENTE_TRANSPORTISTA%>" name="valCd_<%= INTERVINIENTE_TRANSPORTISTA%>" value="<%=auxTransportista %>"/>
                  <input type="hidden" id="valCdOld_<%=INTERVINIENTE_TRANSPORTISTA%>" name="valCdOld_<%=INTERVINIENTE_TRANSPORTISTA%>" value="<%=auxTransportistaOld%>"/>
                  <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_TRANSPORTISTA%>" name="valCuitOld_<%=INTERVINIENTE_TRANSPORTISTA%>" value="<%=auxTransportistaCuitOld %>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT Nro</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_1" name="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TRANSPORTISTA%>);" oncontextmenu="return false" value="<%=auxTransportistaCuit1 %>"/>-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_2" name="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TRANSPORTISTA%>);" oncontextmenu="return false" value="<%=auxTransportistaCuit2 %>"/>-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_3" name="Cuit_<%= INTERVINIENTE_TRANSPORTISTA%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_TRANSPORTISTA%>);" oncontextmenu="return false" value="<%=auxTransportistaCuit3 %>"/>
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_1" name="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_1" value="<%=auxTransportistaCuit1 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_2" name="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_2" value="<%=auxTransportistaCuit2 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_3" name="CuitOld_<%= INTERVINIENTE_TRANSPORTISTA%>_3" value="<%=auxTransportistaCuit3 %>"  />
              </td>
            </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000">CHOFER</td>
              <td align="left">
                  <div id="div_<%= INTERVINIENTE_CHOFER%>" style="display:block;"><%=auxChoferDs%></div>
                  <input type="hidden" id="valDs_<%= INTERVINIENTE_CHOFER%>" name="valDs_<%= INTERVINIENTE_CHOFER%>" value="<%=auxChoferDs %>"/>
                  <input type="hidden" id="valCd_<%= INTERVINIENTE_CHOFER%>" name="valCd_<%= INTERVINIENTE_CHOFER%>" value="<%=auxChoferTipoDoc %>"/>
                  <input type="hidden" id="valCuitOld_<%=INTERVINIENTE_CHOFER%>" name="valCuitOld_<%=INTERVINIENTE_CHOFER%>" value="<%=auxChoferTipoDocOld %>"/>
                  <input type="hidden" id="valCdOld_<%= INTERVINIENTE_CHOFER%>" name="valCdOld_<%= INTERVINIENTE_CHOFER%>" value="<%=auxChoferCuitOld%>"/>
              </td>
              <td width="10%" align="center" style="color:#000">CUIT/CUIL</td>
              <td align="left">
                  <input type="text" size="1" maxlength="2" id="Cuit_<%= INTERVINIENTE_CHOFER%>_1" name="Cuit_<%= INTERVINIENTE_CHOFER%>_1" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CHOFER%>);" oncontextmenu="return false" value="<%=auxChoferNumDoc1 %>" />-
                  <input type="text" size="8" maxlength="8" id="Cuit_<%= INTERVINIENTE_CHOFER%>_2" name="Cuit_<%= INTERVINIENTE_CHOFER%>_2" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CHOFER%>);" oncontextmenu="return false" value="<%=auxChoferNumDoc2 %>" />-
                  <input type="text" size="1" maxlength="1" id="Cuit_<%= INTERVINIENTE_CHOFER%>_3" name="Cuit_<%= INTERVINIENTE_CHOFER%>_3" onkeyPress="return controlDatos(this, event, 'N')" onblur="controlarCUIT(<%= INTERVINIENTE_CHOFER%>);" oncontextmenu="return false" value="<%=auxChoferNumDoc3 %>" />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CHOFER%>_1" name="CuitOld_<%= INTERVINIENTE_CHOFER%>_1" value="<%=auxChoferNumDoc1 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CHOFER%>_2" name="CuitOld_<%= INTERVINIENTE_CHOFER%>_2" value="<%=auxChoferNumDoc2 %>"  />
				  <input type="hidden" id="CuitOld_<%= INTERVINIENTE_CHOFER%>_3" name="CuitOld_<%= INTERVINIENTE_CHOFER%>_3" value="<%=auxChoferNumDoc3 %>"  />
              </td>
            </tr>
	    </tbody>
    </table>
    <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="12" align="center">2</th>
                <td align="left" colspan="8" style="border-top:medium; border-top-right-radius: 8px;" height="18">&nbsp; DATOS DE LOS GRANOS / ESPECIES TRANSPORTADOS</td>
            </tr>
        </thead>
        <tbody>
            <tr>
    	        <td width="12" colspan="2" align="left" style="color:#000000">COSECHA</td>
		        <td width="13%" align="left" style="color:#000000">
                    <input type="text" id="cosecha" name="cosecha" onKeyPress="return controlarDiguitosCosecha(event)" maxlength="8" value="<%= auxCosecha %>"/>
                    <input type="hidden" id="cosechaOld" name="cosechaOld" value="<%= auxCosechaOld %>"/>
                </td>
                <td width="12%" align="left" style="color:#000000">GRANO/ESPECIE</td>
		        <td width="13%" align="left" style="color:#000000">
                   <select id="cmbProducto" name="cmbProducto" onchange="changeProducto(this)">
                   <% Set rsProd = getProductosPto(g_pto) %>
                       <option value="0" <% if(Cdbl(auxGrano) = 0 )then %> selected <% end if %>>Seleccione..</option>
                   <% While not rsProd.Eof %>
                        <option value="<%=rsProd("CDPRODUCTO")%>" <% if(Cdbl(rsProd("CDPRODUCTO")) = Cdbl(auxGrano))then %> selected <% end if %>><%=rsProd("CDPRODUCTO")&"-"&rsProd("DSPRODUCTO") %></option>
                    <%     rsProd.MoveNext()
                      wend %>
                   </select>
                   <input type="hidden" id="cdProductoOld" name="cdProductoOld" value="<%= auxGranoOld %>"/>
                </td>
                <td width="12%" align="left" style="color:#000000">TIPO</td>
		        <td width="13%" align="left" style="color:#000000">&nbsp;</td>
                <td width="12%" align="left" style="color:#000000">CONTRATO Nro</td>
		        <td width="13%" align="left" style="color:#000000">&nbsp;</td>
             </tr>
            <tr>
              <td colspan="3" rowspan="2" align="left" style="color:#000000">CARGA PESADA EN DESTINO
                <input name="" type="radio" value="" /></td>
              <td colspan="2" align="left" style="color:#000000">DECLARACI&OacuteN DE CALIDAD
                <input name="" type="radio" value="" /></td>
              <td align="left" style="color:#000000">BRUTO ORIGEN (Kgrs)</td>
              <td align="left" style="color:#000000">
                <input type="text" id="pesoBruto" name="pesoBruto" onKeyPress="return controlIngreso (this, event, 'N');"  value="<%= auxPesoBruto%>" onblur="calcularPeso(this);"/>
                <input type="hidden" id="pesoBrutoOld" name="pesoBrutoOld" value="<%=auxPesoBrutoOld %>" />
              </td>
              <td colspan="2" align="left" style="color:#000000">CUPO &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
                  <input type="" size="16" id="cupo" name="cupo" value="<%=auxCupo %>" />
                  <input type="hidden" id="cupoOld" name="cupoOld" value="<%=auxCupoOld %>" />
              </td>
              </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000000">CONFORME
                <input name="" type="radio" value="" /></td>
              <td align="left" style="color:#000000">TARA ORIGEN (Kgrs)</td>
              <td align="left" style="color:#000000">
                  <input type="text" id="pesoTara" name="pesoTara" onKeyPress="return controlIngreso (this, event, 'N');"  value="<%= auxPesoTara%>" onblur="calcularPeso(this);"/>
                  <input type="hidden" id="pesoTaraOld" name="pesoTaraOld" value="<%=auxPesoTaraOld %>" />
              </td>
              <td colspan="2" align="left" style="color:#000000">BIOTECNOLOG&IacuteA &nbsp; &nbsp; 
                 <select id="cmbBiotecnologia" name="cmbBiotecnologia" >
                 <option value="0" selected>Seleccione</option>
                 <% Set rsBio = getBiotecnologiaByProducto(g_pto,auxGrano)
                    While not rsBio.Eof %>
                        <option value="<%=rsBio("IDBIOTECNOLOGIA") %>" <% if(Cdbl(rsBio("IDBIOTECNOLOGIA")) = cdbl(auxBiotecnologia))then %> selected <% end if %>><%=rsBio("IDBIOTECNOLOGIA")&"-"&rsBio("DSBIOTECNOLOGIA") %></option>
                 <%    rsBio.MoveNext()
                    wend %>
                 </select>
                 <input type="hidden" id="idBiotecnologiaOld" name="idBiotecnologiaOld" value="<%=auxBiotecnologiaOld %>" />
              </td>
              </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000000">Kgrs. ESTIMADOS</td>
              <td align="left" style="color:#000000">&nbsp;</td>
              <td colspan="2" align="left" style="color:#000000">CONDICIONAL
                <input name="" type="radio" value="" /></td>
              <td align="left" style="color:#000000">NETO ORIGEN (Kgrs)</td>
              <td align="left" style="color:#000000">
                <input type="text" id="pesoNeto" name="pesoNeto" readonly onKeyPress="return controlIngreso (this, event, 'N');"  value="<%= auxPesoNeto %>" onblur="calcularPeso(this);"/>
                <input type="hidden" id="pesoNetoOld" name="pesoNetoOld" value="<%=auxPesoNetoOld %>" />
              </td>
              <td colspan="2" rowspan="4" align="left" style="color:#000000">
                 <textarea name="observaciones" id="observaciones" maxlength="100" rows="5" cols="30"><%=auxObservaciones %></textarea>
                 <input type="hidden" id="observacionesOld" name="observacionesOld" value="<%=auxObservacionesOld %>" />
              </td>
              </tr>
            <tr>
              <td colspan="5" align="center" style="color:#000000;font-weight:bold; ">PROCENDENCIA DE LA MERCADERIA</td>
              <td align="left" style="color:#000000">ESTABLECIMIENTO</td>
              <td align="left" style="color:#000000"><span style="color:#000000; padding:2px"></span></td>
              </tr>
            <tr>
              <td colspan="2" rowspan="2" align="left" style="color:#000000">DIRECCI&OacuteN</td>
              <td colspan="3" rowspan="2" align="left" style="color:#000000">&nbsp;</td>
              <td align="left" style="color:#000000">LOCALIDAD</td>
              <td align="left" style="color:#000000">
                 <input type="" id="procedenciaDs" name="procedenciaDs" value="<%= auxProcedenciaDs %>" />
			     <input type="hidden" name="procedenciaCd" id="procedenciaCd" value="<%= auxProcedenciaCd %>" />
                 <input type="hidden" name="procedenciaCdOld" id="procedenciaCdOld" value="<%= auxProcedenciaCdOld %>" />
              </td>
              </tr>
            <tr>
              <td align="left" style="color:#000000">PROVINCIA</td>
              <td align="left" style="color:#000000">
                <select id="cmbProvincia" name="cmbProvincia" onchange="changeProvincia(this)">  
                <% Set rsProc = getProvinciaProcedencia(g_pto) %>
                <option value="0" <% if(cdbl(auxProcedenciaProv) = 0) then %> selected <% end if %>>Seleccione..</option>
                <% While not rsProc.Eof %>
                      <option value="<%=rsProc("CDPROVINCIA") %>" <% if(Cdbl(rsProc("CDPROVINCIA")) = cdbl(auxProcedenciaProv))then %> selected <% end if %>><%=rsProc("CDPROVINCIA")&"-"&rsProc("DSPROVINCIA") %></option>
                <%    rsProc.MoveNext()
                   wend %> 
                 </select>
                 <input type="hidden" name="cdProvinciaOld" id="cdProvinciaOld" value="<%= auxProcedenciaProvOld %>" />
              </td>
            </tr>
         </tbody>      
      </table>

    <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="10" align="center">3</th>
                <td align="left" colspan="8" style="border-top:medium; border-top-right-radius: 8px;" height="18">&nbsp; LUGAR DE DESTINO DE LOS GRANOS</td>
            </tr>
        </thead>
        <tbody>
            <tr>
              <td colspan="2" rowspan="2" align="left" style="color:#000">DIRECCI&OacuteN</td>
              <td width="40%" colspan="3" rowspan="2" align="left">&nbsp;</td>
              <td align="left" style="color:#000">LOCALIDAD</td>
              <td colspan="6" align="left" width="40%">&nbsp;</td>
            </tr>
            <tr>
              <td align="left" style="color:#000">PROVINCIA</td>
              <td colspan="6" align="left">&nbsp;</td>
            </tr>
        </tbody>      
      </table>

     <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="12" align="center">4</th>
                <td align="left" colspan="8" style="border-top:medium; border-top-right-radius: 8px;" height="18">&nbsp; DATOS DEL TRANSPORTE</td>
            </tr>
        </thead>
        <tbody>
            <tr>
    	    <td width="12" colspan="2" align="left" style="color:#000000">PAGADOR DEL FLETE</td>
		    <td width="13%" align="left" style="color:#000000">&nbsp;</td>
            <td width="12%" align="left" style="color:#000000">CAMI&OacuteN</td>
		    <td width="13%" align="left" style="color:#000000">
                <input type="" id="chapa" name="chapa" value="<%=auxChapa %>" maxlength="6" size="5" style="text-transform:uppercase;"/>
                <input type="hidden" id="chapaOld" name="chapaOld" value="<%=auxChapaOld %>" />
            </td>
            <td width="12%" align="left" style="color:#000000">FLETE PAGO</td>
		    <td width="13%" align="left" style="color:#000000">&nbsp;</td>
            <td width="12%" align="left" style="color:#000000">FLETE A PAGAR</td>
		    <td width="13%" align="left" style="color:#000000">&nbsp;</td>
        </tr>
            <tr>
              <td colspan="2" align="left" style="color:#000000">ACOPLADO</td>
              <td align="left" style="color:#000000">
                <input type="" id="acoplado" name="acoplado" value="<%=auxAcoplado %>" maxlength="6" size="5" style="text-transform:uppercase;"/>
                <input type="hidden" id="acopladoOld" name="acopladoOld" value="<%=auxAcopladoOld %>" />
              </td>
              <td align="left" style="color:#000000">TARIFA DE REFERENCIA</td>
              <td align="left" style="color:#000000">&nbsp;</td>
              <td align="left" style="color:#000000">KMS A RECORRER</td>
              <td align="left" style="color:#000000">&nbsp;</td>
              <td align="left" style="color:#000000">TARIFA</td>
              <td align="left" style="color:#000000">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="5" height="60px" align="center" valign="bottom"><hr width="200" color="#999999">Firma Remitente</td>
              <td colspan="4" height="60px" align="center" valign="bottom"><hr width="200" color="#999999">
                Firma del Chofer</td>
            </tr>
        </tbody>      
      </table>



      <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th width="12" align="center">5</th>
                <td align="left" colspan="8" style="border-top:medium; border-top-right-radius: 8px;" height="18">&nbsp; DATOS A COMPLETAR EN EL LUGAR DE DESTINO Y DESCARGA</td>
            </tr>
        </thead>
        <tbody>
        <tr>
          <td colspan="2" align="left" style="color:#000000">FECHA ARRIBO</td>
          <td width="13%" align="left" style="color:#000000">
            <div id="fechaArriboTxt"><%=auxFechaArribo%></div>
            <input type="hidden" id="fechaArribo" name="fechaArribo" value="<%=auxFechaArribo%>"/>
          </td>
          <td width="13%" align="left" style="color:#000000">HORA</td>
          <td width="12%" align="left" style="color:#000000">              
              <div id="divHoraArribo" ><%=auxHoraArribo %></div>
              <input type="hidden" id="horaArribo" name="horaArribo" value="<%=auxHoraArribo %>"/>
          </td>
          <td width="12%" align="left" style="color:#000000">PESO BRUTO (Kgrs)</td>
          <td width="13%" align="left" style="color:#000000">
             <input type="text" id="pesadaBruto" name="pesadaBruto" value="<%=auxPesadaBruto%>" onKeyPress="return controlIngreso (this, event, 'N');" onblur="calcularPesada()"/>
             <input type="hidden" id="pesadaBrutoOld" name="pesadaBrutoOld" value="<%=auxPesadaBrutoOld %>" />
          </td>
          <td width="25%" colspan="2" rowspan="3" align="left" style="color:#000000">
              <textarea name="observacionesDescarga" id="observacionesDescarga" rows="5" readonly=readonly cols="30"><%=auxObservacionesDescarga %></textarea>
           </td>
        </tr>
        <tr>
          <td colspan="2" align="left" style="color:#000000">FECHA DESCARGA</td>
          <td width="13%" align="left" style="color:#000000">
            <div id="divFechaEgreso"><%=auxFechaEgreso%></div>
            <input type="hidden" id="fechaEgreso" name="fechaEgreso" value="<%=auxFechaEgreso%>"/>
          </td>
          <td width="13%" align="left" style="color:#000000">HORA</td>
          <td width="12%" align="left" style="color:#000000">
             <div id="divHoraDescarga" ><%=auxHoraEgreso %></div>
             <input type="hidden" id="horaDescarga" name="horaDescarga" value="<%=auxHoraEgreso %>" />
          </td>
          <td width="12%" align="left" style="color:#000000">PESO TARA (Kgrs)</td>
          <td width="13%" align="left" style="color:#000000">
              <input type="text" id="pesadaTara" name="pesadaTara" value="<%=auxPesadaTara %>" onKeyPress="return controlIngreso (this, event, 'N');" onblur="calcularPesada()"/>
              <input type="hidden" id="pesadaTaraOld" name="pesadaTaraOld" value="<%=auxPesadaTaraOld %>"/>
          </td>
          </tr>
        <tr>
          <td colspan="2" align="left" style="color:#000000">TURNO NRO</td>
          <td colspan="3" align="left" style="color:#000000">
             <input type="text" id="turno" name="turno" value="<%=auxTurno %>" onKeyPress="return controlIngreso (this, event, 'N');"/>
             <input type="hidden" id="turnoOld" name="turnoOld" value="<%=auxTurnoOld %>" />
          </td>
          <td width="12%" align="left" style="color:#000000">PESO NETO (Kgrs)</td>
          <td width="13%" align="left" style="color:#000000">
             <div id="divPesadaNeto" style="float:right;"><%=GF_EDIT_DECIMALS(auxNetoSMerma,0)%></div>
             <input type="hidden" id="pesadaNeto" name="pesadaNeto" value="<%=auxNetoSMerma%>" />
          </td>
          <input type="hidden" id="mermaKg" name="mermaKg" value="<%=auxMerma%>" />
          <input type="hidden" id="mermaKgOld" name="mermaKgOld" value="<%=auxMermaOld%>" />
          <input type="hidden" id="mermaPorcentaje" name="mermaPorcentaje" value="<%=auxMermaPorcentaje%>" />  
          <input type="hidden" id="mermaPorcentajeOld" name="mermaPorcentajeOld" value="<%=auxMermaPorcentajeOld%>" />
          </tr>
        </tbody>      
      </table>
      <table class="datagrid" width="95%" align="center">
    <thead>
        <tr>
            <th width="10" colspan="1" align="center">6</th>
            <td align="left" colspan="8" style="border-top:medium; border-top-right-radius: 8px;" height="18">&nbsp; CAMBIO DEL DOMICILIO DE DESCARGA / DESVIO</td>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td align="left" colspan="4" style="color:#000">CUIT DESTINO Y DENOMINACI&OacuteN</td>
            <td width="25%" align="left">&nbsp;</td>
            <td width="25%" align="left" style="color:#000">CUIT DESTINATARIO Y DENOMINACI&OacuteN</td>
            <td width="25%" align="left">&nbsp;</td>
        </tr>
        <tr>
          <td align="left" colspan="4" style="color:#000">DOMICILIO</td>
          <td align="left">&nbsp;</td>
          <td align="left" style="color:#000">LOCALIDAD</td>
          <td align="left">&nbsp;</td>
        </tr>
        <tr>
          <td align="left" colspan="4" style="color:#000">NRO PLANTA (ONCCA)</td>
          <td align="left">&nbsp;</td>
          <td colspan="2" rowspan="3" align="center" valign="bottom"><hr width="80%" color="#999999">Firma y sello del Representante o Entregador que ordeno el presente desv&iacuteo</td>
        </tr>
        <tr>
          <td align="left" colspan="4" style="color:#000">FECHA</td>
          <td align="left">&nbsp;</td>
        </tr>
        <tr>
          <td align="left" colspan="4" style="color:#000">TRASLADO ORDENADO POR:</td>
          <td align="left">&nbsp;</td>
        </tr>
    </tbody>      
  </table>
<br/>
   <table width="95%" border="0" align="center" cellpadding="10">
	<tr>
    	<td width="50%">
        <div class="firmas" align="center">
        <table align="center">
        <tr>
        <td align="center"><hr width="180px" color="#999999" style="margin-right:20px">Firma del Perito Recibidor</td>
        <td align="center"><hr width="180px" color="#999999" style="margin-left:20px">Matricula Nro</td>
        </tr>
        </table>
    	</div>
        </td>
        <td width="50%">
        <div class="firmas" align="center">
		<table align="center">
        <tr>
        <td align="center"><hr width="180px" color="#999999" style="margin-right:20px">Firma del Destinatario/Entregador</td>
        <td align="center"><hr width="180px" color="#999999" style="margin-left:20px">Matricula Nro</td>
        </tr>
        </table>
    	</div>
        </td>
    </tr>
</table>


    <input type="hidden" name="accion" id="accion" value="<%=accion%>">
    <input type="hidden" name="runtEvent" id="runtEvent">
    <input type="hidden" name="idCamion" id="idCamion" value="<%=g_IdCamion%>">
    <input type="hidden" name="pto" id="pto" value="<%=g_pto%>">
    <input type="hidden" name="cartaPorteOld" id="cartaPorteOld" value="<%=g_ctaPteOld%>">
</form>	
</body>
</html>