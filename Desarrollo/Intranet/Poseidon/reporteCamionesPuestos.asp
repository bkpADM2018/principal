<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosValidacion.asp"-->
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
Dim strSQL,rs,flagCall,g_fecInicioD,g_fecInicioM,g_fecInicioA,g_fecInicioH,g_fecInicioMi,g_fecFinD,g_fecFinM,g_fecFinA,g_fecFinH,g_fecFinMi,g_accion

Call GP_CONFIGURARMOMENTOS()
pto = GF_PARAMETROS7("pto", "", 6)
Call addParam("pto", pto, params)
g_strPuerto = pto
g_accion	= GF_PARAMETROS7("accion", "", 6)

g_fecInicioD = GF_PARAMETROS7("fecInicioD", 0, 6)
if g_fecInicioD = 0 then g_fecInicioD = GF_nDigits(Day(Now()),2)
g_fecInicioM = GF_PARAMETROS7("fecInicioM", 0, 6)
if g_fecInicioM = 0 then g_fecInicioM = GF_nDigits(Month(Now()),2)
g_fecInicioA = GF_PARAMETROS7("fecInicioA", 0, 6)
if g_fecInicioA = 0 then g_fecInicioA = GF_nDigits(Year(Now()),4)
g_fecInicioH = GF_PARAMETROS7("fecInicioH", 0, 6)
g_fecInicioMi = GF_PARAMETROS7("fecInicioMi", 0, 6)
g_fecInicioS = GF_PARAMETROS7("fecInicioS", 0, 6)
Call GF_STANDARIZAR_MM(g_fecInicioH,g_fecInicioMi,g_fecInicioS)


g_fecFinD = GF_PARAMETROS7("fecFinD", 0, 6)
if g_fecFinD = 0 then g_fecFinD = GF_nDigits(Day(Now()),2)
g_fecFinM = GF_PARAMETROS7("fecFinM", 0, 6)
if g_fecFinM = 0 then g_fecFinM = GF_nDigits(Month(Now()),2)
g_fecFinA = GF_PARAMETROS7("fecFinA", 0, 6)
if g_fecFinA = 0 then g_fecFinA = GF_nDigits(Year(Now()),4)
g_fecFinH = GF_PARAMETROS7("fecFinH", 0, 6)
g_fecFinMi = GF_PARAMETROS7("fecFinMi", 0, 6)
g_fecFinS = GF_PARAMETROS7("fecFinS", 0, 6)
Call GF_STANDARIZAR_MM(g_fecFinH,g_fecFinMi,g_fecFinS)
   

flagCall=false
if (g_accion = ACCION_SUBMITIR) then
    ret = GF_CONTROL_PERIODO(g_fecInicioD, g_fecFinD, g_fecInicioM, g_fecFinM, g_fecInicioA, g_fecFinA)
	Select case (ret)
	    case 0
            if(GF_ControlHora(g_fecInicioH,g_fecInicioMi,g_fecInicioS)) then
                if (GF_ControlHora(g_fecFinH,g_fecFinMi,g_fecFinS)) then
                    g_fechaFin = g_fecFinA & g_fecFinM & g_fecFinD & g_fecFinH & g_fecFinMi & g_fecFinS
                    g_fechaInicio = g_fecInicioA & g_fecInicioM & g_fecInicioD & g_fecInicioH & g_fecInicioMi & g_fecInicioS
                    if (CDbl(g_fechaFin) >= Cdbl(g_fechaInicio)) then
                        flagCall=true
                    else    
                        Call setError(PERIODO_ERRONEO)
                    end if
                else
                    Call setError(FECHA_FIN_INCORRECTA)
                end if
            else
                Call setError(FECHA_INICIO_INCORRECTA)
            end if
        case 1
		    Call setError(FECHA_INICIO_INCORRECTA)
		case 2
		    Call setError(FECHA_FIN_INCORRECTA)
		case 3
		    Call setError(PERIODO_ERRONEO)
    end select
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Puertos - Reporte de Camiones por Puesto</title>
<link rel="stylesheet" type="text/css" href="../css/main.css"> 
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>	
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">

    <% if (flagCall) then %>
        window.open("reporteCamionesPuestosPrintXLS.asp?pto=<%=pto%>&fechaInicio=<%=g_fechaInicio%>&fechaFin=<%=g_fechaFin%>");
    <% end if %>
    
    function bodyOnload() {
		document.forms[0].elements[0].focus();
	}
</script>
</head>

<body onload="bodyOnload()">

<form name="frmSel" id="frmSel" method="post" action="reporteCamionesPuestos.asp">
<div class="tableaside size100">
    <h3> Reporte de Camiones por turnos </h3>
    <div ><% Call showMessages() %></div>
    <div id="searchfilter" class="tableasidecontent">        
		<div class="col66"></div>        
		<div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Inicio:")%> </div>
        <div class="col26">
   			<table>
				<tr>
					<td>
						<input type="text" size="2" maxLength="4" value="<% =g_fecInicioA  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecInicioA"> /
                        <input type="text" size="1" maxLength="2" value="<% =g_fecInicioM  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecInicioM"> /
                        <input type="text" size="1" maxLength="2" value="<% =g_fecInicioD  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecInicioD"> &nbsp;&nbsp;
                        <input type="text" size="1" maxLength="2" value="<% =g_fecInicioH  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecInicioH"> :
                        <input type="text" size="1" maxLength="2" value="<% =g_fecInicioMi %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecInicioMi"> 
					</td>
				</tr>
			</table>
	    </div>
	    <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%> </div>
        <div class="col26">
   			<table>
				<tr>
					<td>
						<input type="text" size="2" maxLength="4" value="<% =g_fecFinA  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecFinA"> /
                        <input type="text" size="1" maxLength="2" value="<% =g_fecFinM  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecFinM"> /
                        <input type="text" size="1" maxLength="2" value="<% =g_fecFinD  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecFinD"> &nbsp;&nbsp;
                        <input type="text" size="1" maxLength="2" value="<% =g_fecFinH  %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecFinH"> :
                        <input type="text" size="1" maxLength="2" value="<% =g_fecFinMi %>" onKeyPress="return controlIngreso (this, event, 'N');"  name="fecFinMi"> 
					</td>
				</tr>
			</table>
	    </div>
        <span style="text-align:center; clear:both; float:left; width:100%"><input type="submit" value="Exportar xls" ></span>
    </div>
</div><!-- END BUSCAR -->
<br>
<div id="actionLabel" class="confirmsj" style="width:80%;visibility:hidden;"></div>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="pto" name="pto" value="<% =pto %>">
</form>
</body>
</html>
