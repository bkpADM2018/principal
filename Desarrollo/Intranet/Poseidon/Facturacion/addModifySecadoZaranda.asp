<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosFechas.asp"-->
<!--#include file="../../Includes/procedimientosTraducir.asp"-->
<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosFacturacionCalidad.asp"-->
<!--#include file="../../Includes/procedimientosSeguridad.asp"-->
<%
    Const PRECIOS_ZARANDA = "ZARANDA"
    Const PRECIOS_SECADO  = "SECADO"

    Const ACCION_MODIFICAR = "modificar"

'--------------------------------------------------------------------------------------------
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

'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA *********************************
'*************************************************************************************
Dim g_strPuerto, lineasTotales, paginaActual, params, myFechaVigencia
Dim rs, strSQL

g_strPuerto = session("TERMINAL_ACTUAL")
call addParam("pto", g_strPuerto, params)

Call initTaskAccessInfo(TASK_POS_MT_ZAR_Y_SEC, session("DIVISION_PUERTO"))
g_cdConcepto = GF_Parametros7("cc",0,6)
Call addParam("cc", g_cdConcepto, params)
paginaActual  = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual = 1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 50

strSQL="SELECT (YEAR(VIGENCIADESDE)*10000 + MONTH(VIGENCIADESDE)*100 + DAY(VIGENCIADESDE)) DTVIGENCIADESDE,"&_
		" PTODESDE,PTOHASTA,CDMONEDA, PRECIO,PRECIOADICIONAL " &_
		" FROM PRECIOSERVICIOS  where CDCONCEPTO= " & g_cdConcepto & " ORDER BY DTVIGENCIADESDE DESC, PTODESDE ASC" 
Call executeQueryDb(g_strPuerto, rs, "OPEN", strSQL)
Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount
%>
<HTML>
<HEAD>
	<TITLE>Precio Acondicionamiento</TITLE>
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />	
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
    <link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">

    <script type="text/javascript" src="../../scripts/controles.js"></script>
    <script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
    <script type="text/javascript" src="../../scripts/paginar.js"></script>
    <script type="text/javascript" src="../../scripts/channel.js"></script>
    <script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
    <script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
    <script type="text/javascript" src="../../scripts/calendar.js"></script>
    <script type="text/javascript" src="../../scripts/calendar-1.js"></script>
    <script type="text/javascript" src="../../scripts/formato.js"></script>
<script language="javascript">	
	
    var ch = new channel();   

    function onLoadPage(){
        <%if (not rs.eof) then%>
            var pgn = new Paginacion("paginacion");
            pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "addModifySecadoZaranda.asp<% =params %>");
        <%end if%>
    }
	
    function CerrarCal(cal) {cal.hide();}
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
    function SeleccionarCalVigencia(cal, date) {
        var str = new String(date);        
        document.getElementById("dtVigenciaD").value = str.substr(0, 2);
        document.getElementById("dtVigenciaM").value = str.substr(3, 2);
        document.getElementById("dtVigenciaA").value = str.substr(6, 4);
		document.getElementById("fechaDtVigenciaDiv").innerHTML = str ;
		
        if (cal) cal.hide();
    }    

    function newPriceClear(){
		document.getElementById("fechaDtVigenciaDiv").innerHTML = "";
        document.getElementById("dtVigenciaD").value = 0 ;
        document.getElementById("dtVigenciaM").value = 0;
        document.getElementById("dtVigenciaA").value = 0;
        document.getElementById("precioB").value = "0" ;                
		document.getElementById("ptoDesde").value = "0" ;
		document.getElementById("ptoHasta").value = "0" ;
		document.getElementById("precioA").value = "0" ;
        document.getElementById("MonedaPesos").checked = true;        
    }

	function control(precio, ptodesde) {
		var ret = true;
		if (PrecioBase = "") PrecioBase = 0;
		if (PrecioBase < 0) {
			alert("El precio no puede ser negativo");
			ret = false;
		}		
<%	if (g_cdConcepto = SERVICIO_ACOND_SECADO) then	%>
		if (ptoDesde <= 0) {										
			alert("Pto desde no puede ser menor a 1");
			ret = false;
		}		
<%	end if		%>		
		return ret;
	}
    function saveNewPrice() {
        var Dia = document.getElementById("dtVigenciaD").value;
        var Mes = document.getElementById("dtVigenciaM").value;
        var Anio = document.getElementById("dtVigenciaA").value;
        var FechaVigencia = Anio+"-"+Mes+"-"+Dia;
        var PrecioBase = document.getElementById("precioB").value;
        var ptoDesde = document.getElementById("ptoDesde").value;
        var ptoHasta = document.getElementById("ptoHasta").value;
		var tipoMoneda = 0;
		if (document.getElementById("MonedaPesos").checked) {
			tipoMoneda = document.getElementById("MonedaPesos").value;
		} else {
			tipoMoneda = document.getElementById("MonedaDolares").value;
		}
		var PrecioAdicional = document.getElementById("precioA").value;

		if (control(PrecioBase, ptoDesde)) {
			var strParameters;
			strParameters = "dtVigencia=" + FechaVigencia + "&PrecioB=" + PrecioBase + "&ptoDesde=" + ptoDesde + "&ptoHasta=" + ptoHasta + "&PrecioA=" + PrecioAdicional + "&cdMoneda=" + tipoMoneda;
			alert("addModifySecadoZarandaAjax.asp<%=params%>&" + strParameters);
			ch.bind("addModifySecadoZarandaAjax.asp<%=params%>&" + strParameters, "saveNewPrice_cb()");
			ch.send();
		}				
    }
	
	function saveNewPrice_cb(){
        var res = ch.response();
        if (res == '<%=RESPUESTA_OK%>') {
            newPriceClear();
            location.reload();
        }
        else{
            document.getElementById("divError").style.display = "block";
            document.getElementById("divError").innerHTML = res;
        }
    }
</script>
</HEAD>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>

<!-- TABLA DE LAS CARPETAS -->
<div class="col66"></div>
<div id="divError" class="reg_Header_Error" style="display:none;"></div>

	<table class="datagrid" align="center" width="90%">
		<thead>
			<tr>
				<th align="center"  ><%=GF_Traducir("Fecha de Vigencia")%></th>
				<th align="center"  ><%=GF_Traducir("Tipo de Moneda")%></th>
				<th align="center"  ><%=GF_Traducir("Precio Base")%></th>
				<th align="center"  ><%=GF_Traducir("Punto Desde")%></th>
				<th align="center"  ><%=GF_Traducir("Punto Hasta")%></th>            
				<th align="center"  ><%=GF_Traducir("Precio Adicional")%></th>
				<th>-</th>
				<th>-</th>				
			</tr>			
		</thead>		
	<% 	Dim indice
		indice = 0    
		if (not rs.eof) then	%>
			<body>
<%			while ((not rs.eof) and (indice < mostrar))
				%>				
				<tr>
					<td align="center"><% =GF_FN2DTE(rs("DTVIGENCIADESDE"))%></td>
					<td align="center"><% =getSimboloMonedaLetras(rs("CDMONEDA")) %></td>
					<td align="center"><%=GF_EDIT_DECIMALS(cdbl(rs("PRECIO"))*100,2) %></td>
					<td align="center"><%=rs("PTODESDE") %></td>
					<td align="center"><%=rs("PTOHASTA") %></td>
					<td align="center"><%=GF_EDIT_DECIMALS(CDBL(rs("PRECIOADICIONAL"))*100,2) %></td>
					<td></td>
					<td></td>
				</tr>
				<%indice = indice + 1
				rs.MoveNext()
			wend 
		Else	%>			
			<tr>
				<td style="text-align:center;" colspan="7" class="reg_header_navdosHL"><%=GF_TRADUCIR("No se encontraron datos")%></td>
			</tr>			
<%		End If%>                               
			<!-- Linea para actualziar/crear nuevo precio -->
			<tr id="newPrice" name="newPrice">
                <td align="center">
                    <table>
                        <tr>
                            <td>
                                <a href="javascript:MostrarCalendario('img_dtVigencia', SeleccionarCalVigencia)">
                                    <img id="img_dtVigencia" src="../../images/calendar-16.png" title="Seleccionar fecha">
                                </a>
                            </td>	
                            <td>
                                <div id="fechaDtVigenciaDiv"></div>
                            </td>	
                        </tr>	
                        <input type="hidden" id="dtVigenciaD" name="dtVigenciaD" value="0">
                        <input type="hidden" id="dtVigenciaM" name="dtVigenciaM" value="0">
                        <input type="hidden" id="dtVigenciaA" name="dtVigenciaA" value="0">
                    </table>
                </td>
                <td align="center">
                    <input type="radio" name="cdMoneda" id="MonedaPesos" value="<%=MONEDA_PESO_NUMERICO%>" checked> Pesos
					<input type="radio" name="cdMoneda" id="MonedaDolares" value="<%=MONEDA_DOLAR_NUMERICO%>"> Dolares                        
                </td>
                <td align="center">
                    <input type="text" id="precioB" value="" onkeypress="return controlIngreso(this,event,'I')" size="11"/>
                </td>
				<td align="center">				
                    <input type="text" id="ptoDesde" value="<% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "0" %>" onkeypress="return controlIngreso(this,event,'N')" size="5" <% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "disabled" %>/>
                </td>
                <td align="center">
                    <input type="text" id="ptoHasta" value="<% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "0" %>" onkeypress="return controlIngreso(this,event,'N')" size="5" <% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "disabled" %>/>
                </td>
                <td align="center">
                    <input type="text" id="precioA" value="<% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "0" %>" onkeypress="return controlIngreso(this,event,'I')" size="11" <% if (g_cdConcepto = SERVICIO_ACOND_ZARANDA) then response.write "disabled" %>/>
                </td>
				<td align="center">
                    <img src="../../images/save-16.png" id="btnSave" title="Guardar" style="cursor:pointer;" onclick="saveNewPrice()"/>
                </td>
                <td align="center">
                    <img src="../images/cancel-16x16.png" id="btnCancel" title="Cancelar" style="cursor:pointer;" onClick="newPriceClear()" />
                </td>
            </tr>
		</tbody>        
        <tfoot>            
            <tr>
                <td colspan="7"><div id="paginacion"></div></td>
            </tr>
        </tfoot>
    </table>
</body>
</html>
