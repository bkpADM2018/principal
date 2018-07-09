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
<title>Camiones - Info</title>
<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

/*----------->Box ADMININSTRACION Control Panel*/
#cell {	
	}
	#cell .boxround {
		padding: 7px;
	}
}
</style>
<script type="text/javascript">
    function abrirConsultaCamiones() {
        window.open('consultaCamiones.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function abrirInformeCamaraVagones() {
        window.open('operativo/AdministracionOperativos.asp?Pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function administracionFacturasCalidad() {
        window.open('Facturacion/administrarFacturas.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function abrirReporteVisteoCalada() {
        window.open('ReportesCalada/reporteVisteosCalada.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function irLaboratorio() {
        window.open('Laboratorio/administrarAnalisis.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function irReporteCalador() {
        window.open('reporteCalador.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function irReporteDescargaEmbarque() {
        window.open('reporteDescargaEmbarques.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
    }
    function abrirCTGNotSend() {
        window.open('informeCTGNoEnviados.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
	}
	function irEnvioArchivo() {
		window.open('exportarDescargasTerceros.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=500,height=500,top=100, left=100, scrollbars=YES,resizable=YES,titlebar=NO');
	}

</script>
</head>
<body>
	<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
	<tr>
	<td width="50%" valign="top">
	<section id="cell">

	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirConsultaCamiones()">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("Consulta de Camiones")%></td>						
				</tr>
				<tr>
					<td class="textoSeccion"><% =GF_TRADUCIR("Administrar los camiones del puerto.")%></td>
				</tr>
			</tbody>
		</table>
	</div>
	<% if isToepfer(session("KCOrganizacion")) then %>
	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirInformeCamaraVagones()">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("Consulta Operativos")%></td>
				</tr>
				<tr>
					<td class="textoSeccion"><% =GF_TRADUCIR("Administrar los operativos de vagones del puerto.")%></td>
				</tr>
			</tbody>
		</table>
	</div>
	<% end if	%>
	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="administracionFacturasCalidad()">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("Facturaci&oacuten Calidad")%> </td>
				</tr>
				<tr>
					<td class="textoSeccion">Consultar las facturas Pendientes/Emitidas por servicios de acondicionamiento.</td>
				</tr>
			</tbody>
		</table>
	</div>
	<% if isToepfer(session("KCOrganizacion")) then %>
	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirCTGNotSend();">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("CTGs No Enviados a AFIP") %></td>						
				</tr>
				<tr>
					<td class="textoSeccion">Consultar el estado de los CTG de los camiones descargados.</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="irLaboratorio();">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("Laboratorio") %></td>						
				</tr>
				<tr>
					<td class="textoSeccion">Consultar los resultados de C&aacutemara y exporta las solicitudes de an&aacutelisis.</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="boxround">
		<table width="480">
			<tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="irEnvioArchivo();">
				<tr>					
					<td class="title_sec_section"><% =GF_TRADUCIR("Enviar Descargas a Bs As.") %></td>
				</tr>
				<tr>
					<td class="textoSeccion">Enviar los archivos de desacarga a la oficina de Bs As para su proceso.</td>
				</tr>
			</tbody>
		</table>
	</div>
	<% end if%>


	</section>   
	</td>
	</tr>
	</table>
</body>
</html>