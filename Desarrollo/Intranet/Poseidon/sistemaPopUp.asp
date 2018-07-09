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
<link rel="stylesheet" href="../css/ActiSAIntra-1.css" type="text/css">
<title>Administracion - Control Panel</title>
<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

	#cell .boxround {
		padding: 7px;
	}
</style>
<script type="text/javascript">
    function abrirAdmUsuarios() {
        window.open('usuarios.asp?Pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO'); 
    }
    function abrirListaUsuarios() {
		window.open('listUsuarios.asp?Pto=<%=g_strPuerto%>','<%=g_strPuerto%>','width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO'); 
	}
	function abrirParrametros() {
	    window.open('consultaParametros.asp?pto=<%=g_strPuerto%>', '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
	}
	function abrirABMProductos() {
	    window.open("producto/productoAdministrar.asp?Pto=<%=g_strPuerto%>", '<%=g_strPuerto%>', "width=1180,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO");
	}	
	function abrirABMProcedencias() {
	    window.open("procedencias/procedenciasAdministrar.asp?Pto=<%=g_strPuerto%>", '<%=g_strPuerto%>', "width=1100,height=500,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO"); 
	}
	
</script>
</head>
<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
<tr>
<td width="50%" valign="top">
<section id="cell">

<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirAdmUsuarios();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Administraci&oacute;n de Usuarios") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Administrar los Usuarios del puerto") %></td>
	        </tr>
        </tbody>
    </table>
</div>
<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirListaUsuarios();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Listado de Usuarios") %></td>						
	        </tr>
	        <tr>
	            <td class="textoSeccion"></td>
	        </tr>
        </tbody>
    </table>
</div>


<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirParrametros();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Parametros") %></td>						
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Consulte los parametros almacendos") %></td>
	        </tr>
        </tbody>
    </table>
</div>

<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirABMProductos();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Adm. de Productos") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Administrar los Productos") %></td>
	        </tr>
        </tbody>
    </table>
</div>

<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirABMProcedencias();">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Adm. de Procedencias") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Administrar las procedencias") %></td>
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