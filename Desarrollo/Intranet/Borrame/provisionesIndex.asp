<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<html>
<head>

<title>Sistema Provisiones</title>

<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />

<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

</head>

<body >
<div id="toolbar"></div>

<br>
<div class="content_list">
    <li>
        <a href="#" onClick="document.location.href='provisionesAdministrarCancelacionAutomatica.asp';">
            <img align="absMiddle" src="images/otadmin-100.png">
            <h3> <% =GF_TRADUCIR("Autorizacion de provisiones desde cancelacion automatica") %>	</h3>
        </a>
    </li>
    <li>
        <a href="#" onClick="document.location.href='provisionesGenerarCancelacionAutomatica.asp';">
            <img align="absMiddle" src="images/download-100.png">
            <h3> <% =GF_TRADUCIR("Generar movimientos de cancelacion automatica") %>	</h3>
        </a>
    </li>
</div>


</body>
</html>