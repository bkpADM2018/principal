<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosTraducir.asp"-->
<!--#include file="../../Includes/procedimientosFacturacionCalidad.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="../../css/main.css" type="text/css">
<link rel="stylesheet" href="../../css/Actisaintra-1.css" type="text/css">
<title>Modificar - Agregar</title>
<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

	#cell .boxround {
		padding: 7px;
	}
</style>
<script type="text/javascript">
    function abrirAdmPrecio(p_tablePrecio) {
        window.open('addModifySecadoZaranda.asp?cc=' + p_tablePrecio, "_blank", 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO', false);
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
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirAdmPrecio('<% =SERVICIO_ACOND_ZARANDA %>');">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Precios Zaranda") %></td>
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Modificar o Agregar los precios de la zaranda") %></td>
	        </tr>
        </tbody>
    </table>
</div>
<div class="boxround">
	<table width="480">
        <tbody onMouseOver="" onMouseOut="" style="cursor: pointer" onClick="abrirAdmPrecio('<% =SERVICIO_ACOND_SECADO %>');">
	        <tr>
	            <td class="title_sec_section"><% =GF_TRADUCIR("Precios Secado") %></td>						
	        </tr>
	        <tr>
	            <td class="textoSeccion"><% =GF_TRADUCIR("Modificar o Agregar los precios del secado") %>></td>
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