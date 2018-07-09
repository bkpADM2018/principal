<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<%

Dim g_AnioHasta,g_MesHasta,g_AnioDesde,g_MesDesde,accion,flagCall,chkDetalle

g_AnioHasta = GF_PARAMETROS7("anioHasta","",6)
if(g_AnioHasta = "")then g_AnioHasta = GF_nDigits(Year(Now()),4)
g_MesHasta = GF_PARAMETROS7("mesHasta","",6)
if(g_MesHasta = "")then g_MesHasta = GF_nDigits(Month(Now()),2)
g_AnioDesde = GF_PARAMETROS7("anioDesde","",6)
if(g_AnioDesde = "")then g_AnioDesde = GF_nDigits(Year(Now()),4)
g_MesDesde = GF_PARAMETROS7("mesDesde","",6)
if(g_MesDesde = "")then g_MesDesde = GF_nDigits(Month(Now()),2)
g_Importe = GF_PARAMETROS7("importe",2,6)
chkDetalle = GF_PARAMETROS7("chkDetalle","",6)
if isFormSubmit() then
	ret = GF_CONTROL_PERIODO("01","01",g_MesDesde,g_MesHasta,g_AnioDesde,g_AnioHasta)
	Select case (ret)
		case 0
			flagCall = true
		case 1
			Call setError(FECHA_INICIO_INCORRECTA)
		case 2
			Call setError(FECHA_FIN_INCORRECTA)
		case 3
			Call setError(PERIODO_ERRONEO)
	end select
end if

%>
<html>
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=9">
    <link rel="stylesheet" type="text/css" href="css/main.css">
	<LINK rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">
	<LINK rel="stylesheet" href="busqueda.css" type="text/css">
	<SCRIPT type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></SCRIPT>
	<script type="text/javascript" src="scripts/controles.js"></script>	
	<SCRIPT type="text/javascript" src="Scripts/jQueryPopUp.js"></SCRIPT>
	<SCRIPT type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></SCRIPT>
	<script type="text/javascript">		
	var puw;
	function bodyOnLoad(){
	<%	if (flagCall) then %>
			generateXLS();
	<%	end if %>	
	}
	function cerrar() {
		puw = getObjPopUp("popupGastosAsociadosAfe");
		puw.hide();
	}
	function generateXLS(){				
		window.open("comprasAFEconsumoPrintXLS.asp?anioDesde=<%=g_AnioDesde%>&mesDesde=<%=g_MesDesde%>&anioHasta=<%=g_AnioHasta%>&mesHasta=<%=g_MesHasta%>&importe=<%=g_Importe%>&verDetalle=<%=chkDetalle%>");		
	}	
	function submitInfo(){
		document.getElementById("formBusqueda").submit;
	}
	function anularImporte(){		
		if (document.getElementById("chkDetalle").checked) {
			document.getElementById("importe").value = 0;
			document.getElementById("importe").disabled = true;
		}
		else{
			document.getElementById("importe").disabled = false;
		}		
	}
</SCRIPT>
</HEAD>
<BODY onLoad="bodyOnLoad()">
	<FORM name="post" id="formBusqueda" name="formBusqueda" action="comprasAFEPopUp.asp">
		<div class="tableaside size100">			
			<div ><% Call showMessages() %></div>			
			<div id="searchfilter" class="tableasidecontent">
				<INPUT type="hidden" name="numeroPagina" id="numeroPagina" value="<%=numeroPagina%>">
				<INPUT type="hidden" name="registrosPorPagina" id="registrosPorPagina" value="<%=registrosPorPagina%>"> 
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Desde:")%> </div>
				<div class="col16">					
					<INPUT type="text" name="mesDesde" id="mesDesde" value="<%=g_MesDesde%>" size="2" maxLength="2" onKeyPress="return controlIngreso(this,event,'N')">/
					<INPUT type="text" name="anioDesde" id="anioDesde" value="<%=g_AnioDesde%>" size="4" maxLength="4" onKeyPress="return controlIngreso(this,event,'N')">
				</div>
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Hasta:")%> </div>
				<div class="col16"> 					
					<INPUT type="text" name="mesHasta" id="mesHasta" value="<%=g_MesHasta%>" size="2" maxLength="2" onKeyPress="return controlIngreso(this,event,'N')">/
					<INPUT type="text" name="anioHasta" id="anioHasta" value="<%=g_AnioHasta%>" size="4" maxLength="4" onKeyPress="return controlIngreso(this,event,'N')">
				</div>
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Importe")%> </div>
				<div class="col16"> 
					<INPUT type="text" id="importe" name="importe" value="<%=g_Importe%>" size="10" onKeyPress="return controlIngreso(this,event,'I')">&nbsp&nbsp u$s
				</div>					
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Ver Detalle")%> </div>
				<div class="col16">					
					<INPUT type="checkbox" id="chkDetalle" name="chkDetalle" onclick="anularImporte();" <%if (Ucase(chkDetalle) = "ON") then Response.Write " CHECKED " %>>					
				</div>
				<span class="btnaction"><INPUT type="submit" value="Exportar XLS" id="buscando" name="buscando" ></span>
				<input type="hidden" id="accion" name="accion" value="<%=ACCION_SUBMITIR%>">
			</div>
		</div>
	</form>	
</body>
<html>
	