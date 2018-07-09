<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->

<%
'----------------------------------------------------------------------------------------------------------------------------------
'           COMIENZO DE LA PAGINA
'----------------------------------------------------------------------------------------------------------------------------------

g_strPuerto = GF_PARAMETROS7("pto", "", 6)

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/Actisaintra-1.css" type="text/css">
<title>Administración de Cupos</title>
<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

	#cell .boxround {
		padding: 7px;
	}
</style>
<script type="text/javascript">
    function abrirReporteCupos() {       
        window.open('reporteCuposAsignados.asp?Pto=<%=g_strPuerto %>', '<%=g_strPuerto %>', 'width=960,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');        
    }
    function abrirCotrolCupoPatente() {
        window.open('controlCuposPatente.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=940,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function abrirLogistica() {
        window.open('cuposAdministrar.asp?pto=<%=g_strPuerto%>&cuitCupeador=<% =session("CuitOrganizacion") %>', '_blank', 'width=1600,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO, location=no');
    }
</script>
</head>
<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
<tr>
<td width="50%" valign="top">
<section id="cell">

<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirLogistica();">
	        <tr>	            
	            <td class="title_sec_section"><% =GF_TRADUCIR("Administraci&oacute;n de Logistica") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion">Administraci&oacute;n de los cupos otorgados por la terminal. </td>
	        </tr>
        </tbody>
    </table>
</div>
<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirReporteCupos();">
	        <tr>	            
	            <td class="title_sec_section"><% =GF_TRADUCIR("Consulta Operativa & Reporte de Cupos") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion">Sigua minuto a minuto el movimiento en la terminal controlando el cumplimiento de los cupos asignados. </td>
	        </tr>
        </tbody>
    </table>
</div>
<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirCotrolCupoPatente();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Control de Cupos y Patentes") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion">Administre la asignación de patentes a los códigos de cupo emitidos.</td>
	        </tr>
        </tbody>
    </table>
</div>

</section>   
</td>
</tr>
</table>
</body>
</html>