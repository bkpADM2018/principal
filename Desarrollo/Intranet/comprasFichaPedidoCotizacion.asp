<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/MD5.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
'Graba los cambios cuando se aprueba o no la Apertura de Sobres
Function grabarAperturaDeSobre(pEstadoApertura, pIdPedido)
    if (pEstadoApertura = ESTADO_ACTIVO) then 
        Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPCTFIRMASAPERTURA_INS", pIdPedido &"||"& Session("MmtoDato") &"||"& Session("Usuario"))
        pct_idEstado = ESTADO_PCT_ABIERTO
    end if
    if (pEstadoApertura = ESTADO_BAJA) then 
        Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPCTFIRMASAPERTURA_DEL", pIdPedido)
        pct_idEstado = ESTADO_PCT_COTIZADO
    end if
    Call grabarHeaderUpdate()
End Function
'---------------------------------------------------------------------------------------------------------------------------------
Call comprasControlAccesoCM(RES_CC)

Dim tab, idPedido, c1, c2, c3, c4, rsCTZ
idPedido = GF_PARAMETROS7("idPedido",0,6)
tab = GF_PARAMETROS7("tab",1,6)
estadoApertura = GF_PARAMETROS7("estado",0,6)
Call initHeader(idPedido)
if (estadoApertura <> 0) then
    Call grabarAperturaDeSobre(estadoApertura, idPedido)
    Response.End
end if
'Se controla si tiene acceso a la información
if (not checkControlPCT()) then	response.redirect "comprasAccesoDenegado.asp"

select case (tab)
	case 1:
		c1="tabbertabdefault"
	case 2:
		c2="tabbertabdefault"
	case 3:
		c3="tabbertabdefault"
	case 4:
		c4="tabbertabdefault"
	case 5:
		c5="tabbertabdefault"
end select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<title>Sistema de Compras - Ficha de Pedido</title>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
		
	var ch = new channel();
	
	function abrirPedido(id) {
		window.open("comprasPedidoCotizacion.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	
	function abrirREMPIC(idpic){
	    window.open("comprasPIC.asp?verRemitos=true&idCotizacionElegida=" + idpic, "_blank");
	}
	
	function anularCTZCallback(pId){
		var myImg = document.getElementById(pId);
		myImg.src="images/1p.gif";
		myImg.onClick="";
	}
	
	function anularCTZ(idCotizacion, idPedido, img){
		if (confirm("Esta seguro que desea anular este Pedido Interno?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("comprasAnularCTZAjax.asp?idCotizacion=" + idCotizacion + "&idPedido=" + idPedido, "anularCTZCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	
	function editarCTZ(idCTZ) {
		document.location.href = "comprasPIC.asp?idCotizacionElegida=" + idCTZ;
	}
		
	
	function cargarContrato(idPedido) {
		window.open('comprasCTCNuevo.asp?idPedido=' + idPedido, "_blank", "location=no,menubar=no,statusbar=no,height=550,width=650",false);
	}
	function abrirContrato(id, estado, canConfirm) {
		if (estado == "<% =ESTADO_CTC_PENDIENTE %>") {
			if (canConfirm) {
				var puw = new winPopUp('popUpCTCConfirm', 'comprasCTCConfirmar.asp?idContrato='+id, 600, 420, 'Confirmar Contrato', 'location.reload()');
			} else {
				window.open('comprasCTCNuevo.asp?CTC_idContrato=' + id, "_blank", "location=no,menubar=no,statusbar=no,height=550,width=650",false);
			}
		} else {
			if (estado != "<% =ESTADO_CTC_CANCELADO %>") {
			 window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
			}
		}
	}
	
	function anularCTCCallback(pId){	
		var resp = ch.response();
		var myImg = document.getElementById(pId);		
		if (resp != "<% =RESPUESTA_OK %>") {
			myImg.src="images/compras/CTZ_cancel-16x16.png";			
			alert(resp);			
		} else {			
			location.reload();
		}
	}
	
	function anularContrato(idContrato, img){
		if (confirm("Esta seguro que desea anular este Contrato?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("comprasCTCAnularContratoAjax.asp?idContrato=" + idContrato, "anularCTCCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	
</script>
<script language="javascript" src="scripts/tabber.js"></script>
</head>
<body>
	<div class="tabber">
		<div class="tabbertab <% =c1 %>" title="<% =GF_TRADUCIR("Pedido")%>"><!--#include file="comprasFichaPCTtab1.asp"--></div>
		<div class="tabbertab <% =c2 %>" title="<% =GF_TRADUCIR("Presupuestos")%>"><!--#include file="comprasFichaPCTtab2.asp"--></div>
		<div class="tabbertab <% =c4 %>" title="<% =GF_TRADUCIR("Pagos")%>"><!--#include file="comprasFichaPCTtab5.asp"--></div>
		<!--
		Para habilitar cuando se cuente con el modulo de firmas electronica
		<div class="tabbertab <% =c4 %>" title="<% =GF_TRADUCIR("Firmas")%>">
		-->
		<!--include file="comprasFichaºPCTtab4.asp"-->
		<!--</div>-->
		
	</div>
</body>