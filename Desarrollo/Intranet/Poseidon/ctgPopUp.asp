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

<title>Administracion de Embarques</title>

<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/Actisaintra-1.css" type="text/css">

<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

	#cell .boxround {
		padding: 7px;
	}
</style>
<script type="text/javascript">
    function abrirReporteCTG() {
        window.open('ReporteCTGEmbarcados/reporteCTGEmbarcados.asp?Pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO'); 
    }
    function abrirAdministracionMuelle() {
        window.open('AdministracionMuelle.asp?Pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }    
</script>
</head>
<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
<tr>
<td width="50%" valign="top">
<section id="cell">

<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirReporteCTG();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Reporte CTGs embarcados") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion">Administre los reportes de los CTG asignados a cada buque. </td>
	        </tr>
        </tbody>
    </table>
</div>
<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirAdministracionMuelle();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Administracion de Buques") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion">Adminsitre la información de las carga realizadas en cada buque.</td>
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