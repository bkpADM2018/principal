<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosCompras.asp"-->
<!--#include file="../../includes/procedimientosTitulos.asp"-->
<!--#include file="../../includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientosLaboratorio.asp"-->
<%
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

'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim fechaDesdeD,fechaDesdeM,fechaDesdeA,fechaHastaD,fechaHastaM,fechaHastaA, mySticker
Dim fileCode,params,myCdProducto, sp_ret
Dim pto

Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
Call addParam("pto", pto, params)

myCdProducto = GF_PARAMETROS7("cdproducto", 0, 6)
Call addParam("cdproducto", myCdProducto, params)

mySticker = GF_PARAMETROS7("sticker", "", 6)
Call addParam("sticker", mySticker, params)

myCartaPorte = GF_PARAMETROS7("cporte", "", 6)
Call addParam("cporte", myCartaPorte, params)
if (myCartaPorte <> 0) then myCartaPorte = GF_nDigits(myCartaPorte,12)

fechaDesde = GF_PARAMETROS7("fdymd", "", 6)
if (fechaDesde = "") then fechaDesde = Left(session("MmtoDato"), 8)
Call addParam("fdymd", fechaDesde, params)

fechaHasta = GF_PARAMETROS7("fhymd", "", 6)
if (fechaHasta = "") then fechaHasta = Left(session("MmtoDato"), 8)
Call addParam("fhymd", fechaHasta, params)

%>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>An�lisis de Laboratorio</TITLE>
	<link rel="stylesheet" href="../../css/ActisaIntra-1.css" type="text/css" />	
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/iwin.css" type="text/css">	
	<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">		
	
	<meta http-equiv="X-UA-Compatible" content="IE=9">
	
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>
</HEAD>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/toolbar.js"></script>
<script type="text/javascript" src="../../scripts/calendar.js"></script>
<script type="text/javascript" src="../../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../../scripts/paginar.js"></script>
<script type="text/javascript" src="../../scripts/iwin.js"></script>
<script language="javascript">	
	var changeFilters = false;

	function submitInfo() {
	    document.getElementById("frmSel").submit();
	}
	
	function generarSolicitudes(){
	    var puw = new PopUpWindow('popUpExportar', 'exportarResultadosPopUp.asp?pto=<% =pto %>', '500', '325', "Exportar Resultados");								
	}

	function onLoadPage(){
	    tb = new Toolbar('toolbar', 6, '../../images/');				
		tb.addButton("refresh-16.png", "Recargar", "submitInfo()");
		tb.addButton("export-16.png", "Exportar archivo", "generarSolicitudes()");		
		tb.draw();
	}
	function CerrarCal(cal) {
	    cal.hide();
	}

	function MostrarCalendario(p_objID, funcSel) {
	    var dte = new Date();
	    var elem = document.getElementById(p_objID);
	    if (calendar != null) calendar.hide();
	    var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
	    cal.setRange(2010, 2099);
	    cal.create();
	    calendar = cal;
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}

	function SeleccionarCalDesde(cal, date) {
	    var str= new String(date);
	    document.getElementById("fd").value = str;
	    document.getElementById("fdymd").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
	    if (cal) cal.hide();
	}

	function SeleccionarCalHasta(cal, date) {
	    var str = new String(date);
	    document.getElementById("fh").value = str;
	    document.getElementById("fhymd").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);	    
	    if (cal) cal.hide();	    	   
	}
	
</script>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>
<form name="frmSel" id="frmSel" action="administrarAnalisis.asp">
<div class="tableaside size100"> <!-- BUSCAR -->
	<h3> Boletines de An&aacute;lisis</h3>
	<div id="searchfilter" class="tableasidecontent">
	    <div class="col16 reg_header_navdos"> Descargas Desde: </div>
        <div class="col16">
            <input type="text" name="fd" id="fd" onClick="javascript:MostrarCalendario('fd', SeleccionarCalDesde)" value="<% =GF_FN2DTE(fechaDesde) %>" size="10">
            <input type="hidden" name="fdymd" id="fdymd" value="<% =fechaDesde %>">
        </div>            
        <div class="col16 reg_header_navdos"> Descargas Hasta: </div>
        <div class="col16">
            <input type="text" name="fh" id="fh" onClick="javascript:MostrarCalendario('fh', SeleccionarCalHasta)" value="<% =GF_FN2DTE(fechaHasta) %>" size="10">
            <input type="hidden" name="fhymd" id="fhymd" value="<% =fechaHasta %>">
        </div>                    	    
		<div class="col16 reg_header_navdos"> Carta Porte: </div>
        <div class="col16">
            <input type="text" id="cporte" maxLength="12"  name="cporte" value="<% =myCartaPorte %>" onKeyPress="return controlIngreso (this, event, 'N');"> 
	    </div>
	    <div class="col16 reg_header_navdos"> Sticker: </div>
        <div class="col16">
            <input type="text" id="sticker" maxLength="9"  name="sticker" value="<% =mySticker %>" onKeyPress="return controlIngreso (this, event, 'N');"> 
	    </div>		    
        <div class="col16 reg_header_navdos"> Producto: </div>
		<div class="col16"> 
			<% strSQL = "SELECT * FROM Productos ORDER BY DSPRODUCTO"
			Call executeQueryDb(pto, rsProducto, "OPEN",strSQL)	 %>
			<select name="cdproducto" id="cdproducto" >
				<option value=""> Todos </option>
					<%while not rsProducto.eof
						mySelected = ""
						if trim(rsProducto("CDPRODUCTO")) = trim(myCdProducto) then mySelected = "SELECTED"%>
						<option value="<%=rsProducto("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProducto("CDPRODUCTO") & "-" & rsProducto("DSPRODUCTO")%></option>
						<%rsProducto.MoveNext()
					 wend%>
			</select>
        </div>		
		<span class="btnaction"><input type="submit" value="Buscar"></span>
		<input type="hidden" id="pto" name="pto" value="<%=pto%>">
	</div>
</div><!-- END BUSCAR -->
</form>
<div class="col16"></div>
    <TABLE class="datagrid" id="TAB1" align="center" width="1022px">	
         <thead>
            <tr>
		        <th align="center" width="146px" >Fecha Descarga</th>
		        <th align="center" width="146px" >Carta Porte</th>
		        <th align="center" width="146px" >Camion/Vagon</th>
		        <th align="center" width="146px">Nro. Sticker</th>
		        <th align="center" width="146px">Producto</th>
		        <th align="center" width="146px" >Fecha An&aacute;lisis</th>
		        <th align="center" width="146px">Nro. Certificado</th>		        		        
	        </tr>			
         </thead>
         <tbody>
            <td colspan="7" class="TDERROR"> EN CONSTRUCCI&Oacute;N</td>   
         </tbody>            
         <tfoot>
            <tr>
                <td colspan="10"><div id="paginacion"></div></td>
            </tr>
        </tfoot>
    </TABLE>	
</body>
</html>
