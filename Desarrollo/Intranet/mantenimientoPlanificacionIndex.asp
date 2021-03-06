<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<% 
Call initAccessInfo(RES_OT_SM)
%>
<html>
<head>

<title>Sistema de Mantenimiento - Planificaciones</title>

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
		var tb = new Toolbar('toolbar', 5, "images/");	
		tb.addButtonHOME("Home", "irA('almacenIndex.asp')");		
		tb.addButtonRETURN("Mantenimiento", "irA('mantenimientoIndex.asp')");				
		tb.draw();		
		pngfix();
	}
	function Encender(pObj){
		pObj.style.color = 'white';
		pObj.style.backgroundImage="url('images/resaltar.png')"
		pObj.style.backgroundRepeat="no-repeat"
	}
	function Apagar(pObj){
		pObj.style.color = 'black';
		pObj.style.backgroundImage="none";
	}	
</script>
</head>

<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>

<br>

<div class="content_list">
    <li>
        <a href="#" onClick="document.location.href='mantenimientoAdministrarPlanificacion.asp';">
            <img align="absMiddle" src="images/calendar-100.png">
            <h3> <% =GF_TRADUCIR("Administración de Planificaciones") %>	</h3>
            <p> <% =GF_TRADUCIR("Administre las planificaciones de mantenimiento.") %> </p>
        </a>
    </li>
    <li>
        <a href="#" onClick="document.location.href='mantenimientoReportes.asp';">
            <img align="absMiddle" src="images/report-100.png">
            <h3> <% =GF_TRADUCIR("Reportes") %> </h3>
            <p> <% =GF_TRADUCIR("Genere todo tipo de reportes referidos a la planificación del mantenimiento de Equipos.") %> </p>
        </a>
    </li>   
</div>

</body>
</html>