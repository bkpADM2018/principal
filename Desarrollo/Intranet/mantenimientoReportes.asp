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
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
    function irA(pLink) {
        location.href = pLink;
    }	
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/");
		tb.addButtonHOME("Home", "irA('mantenimientoIndex.asp')");		
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
    function abrirRptCumplimiento() {        
        var puw = new winPopUp('popupRpt', 'mantenimientoReporteCumplimientoPopUp.asp', '300', '250', 'Reporte de Cumplimiento', '');
    }
</script>
</head>

<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>

<br>

<div class="content_list">
    <li>
        <a href="#" onclick="javascript:abrirRptCumplimiento()">
            <img align="absMiddle" src="images/report-100.png">
            <h3> <% =GF_TRADUCIR("Reporte de Cumplimiento") %>	</h3>
            <p> <% =GF_TRADUCIR("Metrica que permite analizar el trabajo realizado durante el mes.") %> </p>
        </a>
    </li>    
</div>

</body>
</html>