<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<html>
<head>

<title>Sistema de Mantenimiento</title>

<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />


<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
	function irA(pLink) {
		location.href = pLink;
	}
	
	function bodyOnLoad() {
		/*
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.draw();		
		*/
		pngfix();
	}
</script>
</head>

<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>

<br>
<div class="content_list">
    <li>
        <a href="#" onClick="document.location.href='mantenimientoInventarioIndex.asp';">
            <img align="absMiddle" src="images/master-100.png">
            <h3> <% =GF_TRADUCIR("Inventario") %>	</h3>
            <p> <% =GF_TRADUCIR("Administre todos los equipos instalados en la planta.") %> </p>
        </a>
    </li>
    <li>
        <a href="#" onClick="document.location.href='mantenimientoOTIndex.asp';">
            <img align="absMiddle" src="images/ot-100.png">
            <h3> <% =GF_TRADUCIR("Ordenes de Trabajo") %> </h3>
            <p> <% =GF_TRADUCIR("Administre todas las órdenes de trabajo de los equipos.") %> </p>
        </a>
    </li> 
    <li>
        <a href="#" onClick="document.location.href='mantenimientoPlanificacionIndex.asp';">
            <img align="absMiddle" src="images/calendar-100.png">
            <h3> <% =GF_TRADUCIR("Planificación") %> </h3>
            <p> <% =GF_TRADUCIR("Programe las tareas de mantenimiento que se deben realizar sobre los equipos.") %> </p>
        </a>
    </li>   
</div>


</body>
</html>