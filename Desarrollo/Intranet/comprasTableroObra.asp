<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call comprasControlAccesoCM(RES_OBR)

Const FECHA_MINIMA = 0
Const FECHA_MAXIMA = 1
Const RUTA_TABLERO_OBRA = 0

'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************

Dim isAud
Dim idObra, cdMoneda
Dim obraRespCD, obraRespDS
Dim fechaIniObra, fechaFinObra
Dim idArea,idDetalle
Dim presupuesto,presupuestoTotal

idObra = GF_PARAMETROS7("idObra","",6)
cdMoneda = GF_PARAMETROS7("cdMoneda","",6)
if cdMoneda = "" then cdMoneda = MONEDA_DOLAR

Set rsObra = obtenerListaObras(idObra, "", "", "", "")
if (rsObra.eof) then
	response.redirect "comprasAccesoDenegado.asp"
end if
obraRespCD = rsObra("CDRESPONSABLE")
presupuesto = calcularCostoEstimadoObra(cdMoneda,idObra,idarea,iddetalle)
presupuestoTotal = calcularCostoEstimadoObra(cdMoneda,idObra, 0 , 0)

saldo = presupuesto

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<style type="text/css">
	.titleStyle {
		font-weight: bold;
		font-size: 20px;
	}

	.divOculto {
		display: none;
	}

	option.titulo {
	  font-weight: bold;
	}
	.bordeIframe{
		BORDER-BOTTOM: #F4B800 0px solid;
		BORDER-LEFT: #F4B800 0px solid;
		BORDER-TOP: #F4B800 0px solid;
		BORDER-RIGHT: #F4B800 0px solid;
		text-align: center;
		
		-moz-border-radius:5px 5px 5px 5px
	}

	.ocultar{
		display: none;
	}

	.mostrar{
		display: block;
	}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">

	var xco, xcf, yco, ycf;
	
//-------------------------------------------------------------
	function volver() {	
		location.href="comprasObras.asp";
	}
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	
	function irAdministracion() {
		location.href = "comprasAdministracion.asp";
	}
	
	function recargar() {		
		window.location.reload();
	}
	
	function cambiarMoneda() {		
		document.getElementById("frmMoneda").submit();
	}
	
	function submitir() {		
		document.getElementById("frmMoneda").submit();
	}
	
	function createBudget() {
		window.open('comprasBudgetObra.asp?idObra=<% =idObra %>');
		var img = document.getElementById("imgBudget")
		img.src="images/compras/refresh-16x16.png";
		if (isFirefox) {
			img.setAttribute('onclick', "reloadPage()");
		} else {
			img['onclick'] = new Function("reloadPage()");
		}	
	}
	
	function imprimir(){		
		window.open("comprasbudgetobrafilter.asp?idobra=<%=idObra%>");
	}

	function bodyOnLoad() {	
	
		var tb = new Toolbar('toolbar', 6, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");
		tb.addButtonREFRESH("Recargar", "recargar()");		
		tb.addButton("printer-16x16.png", "Imprimir", "imprimir()");
		tb.addButton("previous-16x16.png", "Atrás", "volver()");
		tb.draw();		
		
		pngfix();

		document.getElementById("detalle").src = "comprasTableroObraDetalle.asp?idobra=<%=idObra%>&cdMoneda=<%=cdMoneda%>&ruta=<%=RUTA_TABLERO_OBRA%>";
	}
	
	function abrirAFEPrint(id){
		window.open("comprasAFEPrint.asp?idAFE=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	
	function iFrameOnLoad()
	{
		reSize();	
		document.getElementById('loading').className = "ocultar";
	}
	
	function reSize()
	{
		try
		{
			var oBody = detalle.document.body;
			var oFrame = document.getElementById("detalle");
			oFrame.style.height = oBody.scrollHeight;
		}
		//An error is raised if the IFrame domain != its container's domain
		catch(e)
		{
			window.status = 'Error: ' + e.number + '; ' + e.description;
		}
		
	}
  </script>

</head>
<body onload="bodyOnLoad();">	
		
	<div id="toolbar"></div>
	<br>
	<form id="frmMoneda" name="frmMoneda">
	<table width="90%" align="center" border="0">
		<tr>
			<td width="70%">&nbsp;</td>
			<td>
				<table width="100%" align="right" cellpadding="2" cellspacing="1" class="reg_Header" border="0">
					<tr>
						<td><% =GF_TRADUCIR("Seleccione Moneda") %>:</td>
						<td>
							<select id="cdMoneda" name="cdMoneda" onChange="cambiarMoneda();">					
								<option value="<%=MONEDA_PESO%>" <%if cdMoneda = MONEDA_PESO then response.write "selected"%> ><% =GF_TRADUCIR("Peso argentino") %></option>					
								<option value="<%=MONEDA_DOLAR%>" <%if cdMoneda = MONEDA_DOLAR then response.write "selected"%> ><% =GF_TRADUCIR("Dolar estadounidense") %></option>
							</select>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>	
	<input id="idObra" name="idObra" type="hidden" value="<%=idObra%>">

	<br>	
	<table width="90%" align="center" cellpadding="2" cellspacing="1" class="reg_Header" border="0">	
		<tr>
			<td width="10%" class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Codigo") %></td>
			<td width="300px" class="reg_Header_navdos">&nbsp;<b><% =rsObra("cdObra") %></b></td> 		
			<td width="550px" class="reg_header_nav"><% =GF_TRADUCIR("AFEs") %></td>
			<td width="10px">&nbsp;</td>
			<td rowSpan="10" valign="middle" align="center">
					&nbsp;
            </td>
		</tr>	

		<tr>
			<td class="reg_Header_nav"><% =GF_TRADUCIR("Descripcion") %></td>					
			<td class="reg_Header_navdos">&nbsp;<b><% =rsObra("dsObra") %></b></td>				
			<td rowSpan="7" valign="top"><!--#include file="comprasListaAFE.asp"--></td>
		</tr>
		<tr>
		
		<td class="reg_Header_nav"><% =GF_TRADUCIR("Responsable") %></td>
			
		<td class="reg_Header_navdos">&nbsp;<b>
				<%obraRespDS = getUserDescription(obraRespCD)
				response.write obraRespDS %></b>			</td>					
		</tr>
		<tr>				
			<td class="reg_Header_nav"><% =GF_TRADUCIR("Fecha Inicio") %></td>
			<td class="reg_Header_navdos">&nbsp;<b><% =GF_FN2DTE(rsObra("FECHAINICIO"))%></b></td>					
		</tr>
		<tr>
			<td class="reg_Header_nav"><% =GF_TRADUCIR("Fecha Fin") %></td>
			<td class="reg_Header_navdos">&nbsp;<b><% =GF_FN2DTE(rsObra("FECHAFIN"))%></b></td>					
		</tr>	
		<tr>
			<td class="reg_Header_nav"><% =GF_TRADUCIR("Fecha Ajustada") %></td>			
			<td class="reg_Header_navdos">&nbsp;<b><% if (rsObra("FECHAAJUSTADA") <> "") then	Response.write GF_FN2DTE(rsObra("FECHAAJUSTADA")) %></b></td>					
		</tr>
		<tr>
			<td class="reg_Header_nav"><% =GF_TRADUCIR("Presupuesto") %></td>						
			<td class="reg_Header_navdos">
				<%	isAud = isAuditor(rsObra("IDDIVISION"))
					if (checkControlObra(rsObra("IDOBRA")) or (isAud)) then%>
						<% =getSimboloMoneda(cdMoneda) & " " & GF_EDIT_DECIMALS(presupuestoTotal,2) %></b>
						<%  if ((not isAud) and (puedeModificarBudget(rsObra("CDRESPONSABLE"), rsObra("FECHAINICIO"), rsObra("IDDIVISION")))) then	 %>				
								<img style="cursor:pointer" id="imgBudget<% =rsObra("IDOBRA") %>" src="images/compras/edit-16x16.png" title="<% =GF_TRADUCIR("Cargar/Modificar Presupuesto") %>" onClick="javascript:createBudget(<% =rsObra("IDOBRA") %>)">
						<%	else
							if (puedeReasignarBudget(rsObra("CDRESPONSABLE"), rsObra("IDDIVISION"))) then%>
								<a onClick="javascript:location.href='comprasBudgetReasignaciones.asp?idObra=<% =rsObra("IDOBRA") %>' "><img style="cursor:pointer" id="imgBudget<% =rsObra("IDOBRA") %>" src="images/compras/budget_view-16x16.png" title="<% =GF_TRADUCIR("Reasignar Presupuesto") %>"></a>
							<%else%>
								<a onClick="javascript:window.open('comprasbudgetobrafilter.asp?idObra=<% =rsObra("IDOBRA") %>')"><img style="cursor:pointer" id="imgBudget<% =rsObra("IDOBRA") %>" src="images/compras/budget_view-16x16.png" title="<% =GF_TRADUCIR("Ver Detalle Presupuesto") %>"></a>
							<%end if
						end if	
					end if	%>			
			</td>					
		</tr>		
		<tr>
			<td class="reg_Header_nav"><% =GF_TRADUCIR("División") %></td>
			<td class="reg_Header_navdos"><% =getDescripcionDivision(rsObra("IDDIVISION"))%></b></td>
		</tr>
		<tr>
		  <td class="reg_Header_nav round_border_bottom_left"><% =GF_TRADUCIR("Fotos") %></td>
		  <td class="reg_Header_navdos">&nbsp;<a href="comprasObrasFotos.asp?idObra=<%=idObra%>"> <img src="images/compras/Picture-icon-16x16.png" border="0"></a> </td>
		  <td>&nbsp;</td>
		</tr>									
	</table>
	<br />
	<table width="100%" border="0">
		<tr>
			<td align="center">
				<img src="images/compras/loading_big.gif" id="loading" name="loading" class="">
				<iframe width="90%" height="500px" class="bordeIframe mostrar" id="detalle" name="detalle"  ></iframe>
			</td>
		</tr>
	</table>	
</form>
</body>
</html>